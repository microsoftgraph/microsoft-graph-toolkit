// A workaround to load icon font inside Shadow Dom.
const font = document.createElement('link');
font.rel = 'stylesheet';
font.href = '/src/icons/icon-font.css';
document.head!.appendChild(font);
