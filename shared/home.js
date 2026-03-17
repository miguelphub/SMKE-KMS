document.addEventListener("DOMContentLoaded", () => {
  const particleHost = document.getElementById("bgParticles");

  if (particleHost) {
    const particleCount = window.innerWidth < 640 ? 22 : 38;
    const fragment = document.createDocumentFragment();

    for (let i = 0; i < particleCount; i += 1) {
      const particle = document.createElement("span");
      particle.style.left = `${Math.random() * 100}%`;
      particle.style.top = `${Math.random() * 100}%`;
      particle.style.animationDuration = `${7 + Math.random() * 8}s`;
      particle.style.animationDelay = `${Math.random() * 5}s`;
      particle.style.opacity = `${0.18 + Math.random() * 0.42}`;
      particle.style.transform = `scale(${0.8 + Math.random() * 1.3})`;
      fragment.appendChild(particle);
    }

    particleHost.appendChild(fragment);
  }

  document.querySelectorAll(".nav-card").forEach((card, index) => {
    card.style.setProperty("--delay", `${index * 90}ms`);
  });
});
