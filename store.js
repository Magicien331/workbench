// Data layer. Two backends with the same interface:
//  - Firebase (Google sign-in + Firestore), used when firebase-config.js is filled in
//  - Demo (this browser only), used before Firebase is set up and for testing
import { FIREBASE_CONFIG } from './firebase-config.js';

export const isConfigured = !!(FIREBASE_CONFIG && FIREBASE_CONFIG.apiKey && !/^PASTE/i.test(FIREBASE_CONFIG.apiKey));

export async function createStore(){
  return isConfigured ? firebaseStore() : demoStore();
}

const lower = s=>String(s||'').trim().toLowerCase();
const randomId = (n=20)=>{ const a = new Uint8Array(n), c = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'; crypto.getRandomValues(a); return [...a].map(x=>c[x % 62]).join(''); };
const userOf = u=>u && {uid:u.uid, email:lower(u.email), name:u.displayName || u.email || '',
  photo:u.photoURL || (u.providerData || []).map(p=>p && p.photoURL).find(Boolean) || ''}; // Google sometimes leaves the top-level photo empty

/* ======================= Firebase ======================= */
async function firebaseStore(){
  const [A0, AU, F] = await Promise.all([
    import('./vendor/firebase/firebase-app.js'),
    import('./vendor/firebase/firebase-auth.js'),
    import('./vendor/firebase/firebase-firestore.js')
  ]);
  const app = A0.initializeApp(FIREBASE_CONFIG);
  const auth = AU.getAuth(app);
  let db;
  try { db = F.initializeFirestore(app, {localCache: F.persistentLocalCache({tabManager: F.persistentMultipleTabManager()})}); }
  catch(e){ db = F.getFirestore(app); }
  const col = (...p)=>F.collection(db, ...p);
  const ref = (...p)=>F.doc(db, ...p);
  const docs = snap=>snap.docs.map(d=>({id:d.id, data:d.data()}));

  return {
    mode:'firebase',
    onAuth(cb, onErr){ AU.getRedirectResult(auth).catch(e=>onErr?.(e)); return AU.onAuthStateChanged(auth, u=>cb(userOf(u))); },
    async signIn(){
      const p = new AU.GoogleAuthProvider(); p.setCustomParameters({prompt:'select_account'});
      try { await AU.signInWithPopup(auth, p); }
      catch(e){
        if (['auth/popup-blocked','auth/operation-not-supported-in-this-environment'].includes(e.code)) await AU.signInWithRedirect(auth, p);
        else if (!['auth/popup-closed-by-user','auth/cancelled-popup-request'].includes(e.code)) throw e;
      }
    },
    // Sign out and wipe this device's offline copy of the data
    async signOut(){
      await AU.signOut(auth);
      try { await F.terminate(db); await F.clearIndexedDbPersistence(db); } catch(e){ console.warn(e); }
      location.reload();
    },
    // Only people listed in the "allow" collection (added by the site owner in the Firebase console) can create workspaces
    async canCreate(email){ try { return (await F.getDoc(ref('allow', email))).exists(); } catch(e){ return false; } },

    // workspaces the user owns or was given access to
    watchWorkspaces(user, cb, onErr){
      let owned = null, shared = null;
      const emit = ()=>{ if (owned && shared){ const m = new Map(); [...owned, ...shared].forEach(w=>m.set(w.id, w)); cb([...m.values()]); } };
      // An empty answer from the offline cache isn't trusted (it would show the welcome screen to someone who has boards)
      const take = s=>!(s.metadata.fromCache && s.empty);
      const u1 = F.onSnapshot(F.query(col('workspaces'), F.where('ownerUid','==',user.uid)), {includeMetadataChanges:true},
        s=>{ if (!take(s) && owned === null) return; owned = docs(s).map(d=>({id:d.id, ...d.data})); emit(); }, e=>{ owned = owned || []; emit(); onErr?.(e); });
      const u2 = F.onSnapshot(F.query(col('workspaces'), F.where('memberEmails','array-contains',user.email)), {includeMetadataChanges:true},
        s=>{ if (!take(s) && shared === null) return; shared = docs(s).map(d=>({id:d.id, ...d.data})); emit(); }, e=>{ shared = shared || []; emit(); onErr?.(e); });
      return ()=>{ u1(); u2(); };
    },
    async createWorkspace(data){ const r = ref('workspaces', data.ownerUid + '_' + randomId()); await F.setDoc(r, data); return r.id; },
    updateWorkspace(id, patch){ return F.updateDoc(ref('workspaces', id), patch); },
    setMember(id, email, role){
      const r = ref('workspaces', id), fp = new F.FieldPath('roles', email);
      return role ? F.updateDoc(r, fp, role, 'memberEmails', F.arrayUnion(email))
                  : F.updateDoc(r, fp, F.deleteField(), 'memberEmails', F.arrayRemove(email));
    },

    // boards, issues and settings of one workspace
    watchData(wsId, {boards, issues, meta}, onErr){
      const a = F.onSnapshot(col('workspaces', wsId, 'boards'), s=>boards(docs(s)), onErr);
      const b = F.onSnapshot(col('workspaces', wsId, 'issues'), s=>issues(docs(s)), onErr);
      const c = F.onSnapshot(ref('workspaces', wsId, 'meta', 'settings'), s=>meta(s.exists() ? s.data() : null), onErr);
      return ()=>{ a(); b(); c(); };
    },
    // campaigns: small documents named "cmp_…" kept next to the settings
    watchCampaigns(wsId, cb, onErr){
      const q = F.query(col('workspaces', wsId, 'meta'), F.where(F.documentId(), '>=', 'cmp_'), F.where(F.documentId(), '<', 'cmp`'));
      return F.onSnapshot(q, s=>cb(docs(s)), onErr);
    },
    async commit(wsId, ops){
      const chunks = []; let cur = [], size = 0;
      for (const o of ops){
        const n = o.data ? JSON.stringify(o.data).length : 100;
        // small batches keep each one under the security rules' limit on document lookups
        if (cur.length && (cur.length >= 10 || size + n > 6e6)){ chunks.push(cur); cur = []; size = 0; }
        cur.push(o); size += n;
      }
      if (cur.length) chunks.push(cur);
      for (const part of chunks){
        const batch = F.writeBatch(db);
        part.forEach(o=>{ const r = ref('workspaces', wsId, o.col, o.id); o.type === 'del' ? batch.delete(r) : batch.set(r, o.data); });
        await batch.commit();
      }
    },

    // the activity history: newest "act_…" documents first (their names sort by time)
    async listActivity(wsId, n=500){
      const q = F.query(col('workspaces', wsId, 'meta'), F.where(F.documentId(), '>=', 'act_'), F.where(F.documentId(), '<', 'act`'), F.orderBy(F.documentId(), 'desc'), F.limit(n));
      return docs(await F.getDocs(q)).map(d=>({id:d.id, ...d.data}));
    },
    // access requests
    async createRequest(data){
      const r = ref('requests', data.fromUid + '_' + data.toEmail);
      await F.deleteDoc(r).catch(()=>{}); // an earlier declined request to the same person
      await F.setDoc(r, data);
    },
    watchIncoming(email, cb, onErr){ return F.onSnapshot(F.query(col('requests'), F.where('toEmail','==',email)), s=>cb(docs(s).map(d=>({id:d.id, ...d.data}))), onErr); },
    watchOutgoing(uid, cb, onErr){ return F.onSnapshot(F.query(col('requests'), F.where('fromUid','==',uid)), s=>cb(docs(s).map(d=>({id:d.id, ...d.data}))), onErr); },
    updateRequest(id, patch){ return F.updateDoc(ref('requests', id), patch); },
    deleteRequest(id){ return F.deleteDoc(ref('requests', id)); },
    // one private document per person per workspace (only that person can read or write it)
    async getPrivate(wsId, uid){ const s = await F.getDoc(ref('workspaces', wsId, 'private', uid)); return s.exists() ? s.data() : null; },
    setPrivate(wsId, uid, data){ return F.setDoc(ref('workspaces', wsId, 'private', uid), data); }
  };
}

/* ======================= Demo backend (this browser only) ======================= */
function demoStore(){
  const KEY = 'workbench.demo.db', UKEY = 'workbench.demo.user';
  const load = ()=>{ try { return JSON.parse(localStorage.getItem(KEY)) || null; } catch(e){ return null; } };
  let db = load() || {workspaces:{}, data:{}, requests:{}};
  const subs = new Set();
  const persist = ()=>{ try { localStorage.setItem(KEY, JSON.stringify(db)); } catch(e){} fire(); };
  const fire = ()=>queueMicrotask(()=>subs.forEach(f=>f()));
  addEventListener('storage', e=>{ if (e.key === KEY){ db = load() || db; fire(); } if (e.key === UKEY) authFire(); });
  const sub = fn=>{ subs.add(fn); queueMicrotask(fn); return ()=>subs.delete(fn); };
  const uid = ()=>Math.random().toString(36).slice(2,12);
  const clone = o=>JSON.parse(JSON.stringify(o));
  const authCbs = new Set();
  const getUser = ()=>{ try { return JSON.parse(sessionStorage.getItem(UKEY)) || null; } catch(e){ return null; } };
  const authFire = ()=>authCbs.forEach(cb=>cb(getUser()));
  const me = ()=>getUser();
  const canRead = (w, u)=>u && (w.ownerUid === u.uid || (w.memberEmails||[]).includes(u.email));
  const canWrite = (w, u)=>u && (w.ownerUid === u.uid || (w.roles||{})[u.email] === 'editor');
  const deny = ()=>Promise.reject(Object.assign(new Error('Missing or insufficient permissions.'), {code:'permission-denied'}));

  return {
    mode:'demo',
    onAuth(cb, onErr){ authCbs.add(cb); queueMicrotask(()=>cb(getUser())); return ()=>authCbs.delete(cb); },
    async signIn(profile){
      // profile comes from the demo sign-in form: {name, email}
      const email = lower(profile?.email); if (!email) throw new Error('email required');
      const u = {uid:'demo-' + email.replace(/[^a-z0-9]/g,'_'), email, name:profile.name || email.split('@')[0], photo:''};
      sessionStorage.setItem(UKEY, JSON.stringify(u)); authFire();
    },
    async signOut(){ sessionStorage.removeItem(UKEY); authFire(); },
    async canCreate(){ return true; },

    watchWorkspaces(user, cb){
      return sub(()=>cb(Object.entries(db.workspaces).filter(([,w])=>canRead(w, user)).map(([id,w])=>({id, ...clone(w)}))));
    },
    async createWorkspace(data){ const id = uid(); db.workspaces[id] = clone(data); db.data[id] = {boards:{}, issues:{}, meta:null}; persist(); return id; },
    async updateWorkspace(id, patch){
      const w = db.workspaces[id]; if (!w || w.ownerUid !== me()?.uid) return deny();
      Object.assign(w, clone(patch)); persist();
    },
    async setMember(id, email, role){
      const w = db.workspaces[id]; if (!w || w.ownerUid !== me()?.uid) return deny();
      w.roles = w.roles || {}; w.memberEmails = w.memberEmails || [];
      if (role){ w.roles[email] = role; if (!w.memberEmails.includes(email)) w.memberEmails.push(email); }
      else { delete w.roles[email]; w.memberEmails = w.memberEmails.filter(e=>e !== email); }
      persist();
    },
    watchData(wsId, {boards, issues, meta}){
      return sub(()=>{
        const w = db.workspaces[wsId], d = db.data[wsId];
        if (!w || !d || !canRead(w, me())) return;
        boards(Object.entries(d.boards).map(([id,data])=>({id, data:clone(data)})));
        issues(Object.entries(d.issues).map(([id,data])=>({id, data:clone(data)})));
        meta(d.meta ? clone(d.meta) : null);
      });
    },
    async commit(wsId, ops){
      const w = db.workspaces[wsId]; if (!w || !canWrite(w, me())) return deny();
      const d = db.data[wsId];
      ops.forEach(o=>{
        if (o.col === 'meta' && o.id.startsWith('cmp_')){ d.camps = d.camps || {}; if (o.type === 'del') delete d.camps[o.id]; else d.camps[o.id] = clone(o.data); return; }
        if (o.col === 'meta' && o.id !== 'settings'){ d.act = d.act || {}; if (o.type === 'del') delete d.act[o.id]; else d.act[o.id] = clone(o.data); return; }
        if (o.col === 'meta'){ d.meta = o.type === 'del' ? null : clone(o.data); return; }
        if (o.type === 'del') delete d[o.col][o.id]; else d[o.col][o.id] = clone(o.data);
      });
      persist();
    },
    watchCampaigns(wsId, cb){
      return sub(()=>{ const w = db.workspaces[wsId], d = db.data[wsId]; if (!w || !d || !canRead(w, me())) return; cb(Object.entries(d.camps || {}).map(([id, data])=>({id, data:clone(data)}))); });
    },
    async listActivity(wsId, n=500){ const d = db.data[wsId]; if (!d || !canRead(db.workspaces[wsId], me())) return deny(); return Object.entries(d.act || {}).sort((a, b)=>a[0] < b[0] ? 1 : -1).slice(0, n).map(([id, x])=>({id, ...clone(x)})); },
    async createRequest(data){ const id = data.fromUid + '_' + data.toEmail; delete db.requests[id]; db.requests[id] = clone(data); persist(); },
    watchIncoming(email, cb){ return sub(()=>cb(Object.entries(db.requests).filter(([,r])=>r.toEmail === email).map(([id,r])=>({id, ...clone(r)})))); },
    watchOutgoing(u, cb){ return sub(()=>cb(Object.entries(db.requests).filter(([,r])=>r.fromUid === u).map(([id,r])=>({id, ...clone(r)})))); },
    async updateRequest(id, patch){ const r = db.requests[id]; if (!r || r.toEmail !== me()?.email) return deny(); Object.assign(r, clone(patch)); persist(); },
    async deleteRequest(id){ delete db.requests[id]; persist(); },
    async getPrivate(wsId, u){ return clone((db.private || {})[wsId + '/' + u] ?? null); },
    async setPrivate(wsId, u, data){ if (!me() || me().uid !== u) return deny(); (db.private ||= {})[wsId + '/' + u] = clone(data); try { localStorage.setItem(KEY, JSON.stringify(db)); } catch(e){} }
  };
}
