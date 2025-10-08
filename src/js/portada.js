
document.addEventListener('DOMContentLoaded', () => {
  const y = document.querySelector('[data-year]');
  if(y) y.textContent = new Date().getFullYear();
});
