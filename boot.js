// Refuse to run inside someone else's frame (clickjacking)
if (window.top !== window.self){ document.documentElement.hidden = true; try { window.top.location = window.location.href; } catch (e) {} }
// Set language, direction and theme before anything paints
try {
  var l = localStorage.getItem('workbench.lang') || 'he';
  document.documentElement.lang = l; document.documentElement.dir = l === 'he' ? 'rtl' : 'ltr';
  var p = JSON.parse(localStorage.getItem('workbench.prefs') || '{}');
  var th = p.theme || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'); document.documentElement.dataset.theme = th; document.documentElement.dataset.tone = ['dark','purpledark','blackred','blackyellow'].indexOf(th) >= 0 ? 'dark' : 'light';
} catch (e) {}
