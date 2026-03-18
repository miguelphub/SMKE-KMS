document.addEventListener("DOMContentLoaded", () => {
  const particleHost = document.getElementById("bgParticles");

  if (particleHost) {
    const particleCount = window.innerWidth < 640 ? 10 : 16;
    const fragment = document.createDocumentFragment();

    for (let i = 0; i < particleCount; i += 1) {
      const particle = document.createElement("span");
      particle.style.left = `${Math.random() * 100}%`;
      particle.style.top = `${Math.random() * 100}%`;
      particle.style.animationDuration = `${10 + Math.random() * 8}s`;
      particle.style.animationDelay = `${Math.random() * 5}s`;
      particle.style.opacity = `${0.04 + Math.random() * 0.12}`;
      fragment.appendChild(particle);
    }

    particleHost.appendChild(fragment);
  }

  document.querySelectorAll(".nav-card").forEach((card, index) => {
    card.style.setProperty("--delay", `${index * 70}ms`);
  });
});
