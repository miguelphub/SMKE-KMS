(function(){
  const KEY = 'kms-theme';
  const root = document.documentElement;
  let switchTimer = null;

  function markThemeTransition(){
    root.classList.add('theme-switching');
    window.clearTimeout(switchTimer);
    switchTimer = window.setTimeout(() => {
      root.classList.remove('theme-switching');
    }, 430);
  }

  function applyTheme(theme){
    root.dataset.theme = theme;
    document.querySelectorAll('[data-theme-toggle]').forEach((button) => {
      const icon = button.querySelector('i');
      const label = button.querySelector('[data-theme-label]');
      if (theme === 'dark') {
        if (icon) icon.className = 'fa-solid fa-sun';
        if (label) label.textContent = 'Modo claro';
        button.setAttribute('aria-label', 'Cambiar a modo claro');
      } else {
        if (icon) icon.className = 'fa-solid fa-moon';
        if (label) label.textContent = 'Modo oscuro';
        button.setAttribute('aria-label', 'Cambiar a modo oscuro');
      }
    });
  }

  function nextTheme(){
    return (root.dataset.theme || 'light') === 'dark' ? 'light' : 'dark';
  }

  document.addEventListener('DOMContentLoaded', () => {
    applyTheme(localStorage.getItem(KEY) || root.dataset.theme || 'light');
    document.querySelectorAll('[data-theme-toggle]').forEach((button) => {
      button.addEventListener('click', () => {
        markThemeTransition();
        const theme = nextTheme();
        localStorage.setItem(KEY, theme);
        applyTheme(theme);
      });
    });
  });
})();
