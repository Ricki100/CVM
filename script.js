const header = document.querySelector("[data-header]");
const progress = document.querySelector(".scroll-progress");
const revealItems = document.querySelectorAll(".reveal");
const parallaxItems = document.querySelectorAll("[data-parallax]");
const primaryNav = header?.querySelector(".nav-links");

if (header && primaryNav) {
  const menuToggle = document.createElement("button");
  menuToggle.className = "menu-toggle";
  menuToggle.type = "button";
  menuToggle.setAttribute("aria-label", "Open navigation menu");
  menuToggle.setAttribute("aria-expanded", "false");
  menuToggle.innerHTML = "<span aria-hidden=\"true\"></span>";
  header.insertBefore(menuToggle, header.querySelector(".nav-cta"));

  function setMenuOpen(isOpen) {
    header.classList.toggle("is-menu-open", isOpen);
    document.body.classList.toggle("is-menu-open", isOpen);
    menuToggle.setAttribute("aria-expanded", String(isOpen));
    menuToggle.setAttribute("aria-label", isOpen ? "Close navigation menu" : "Open navigation menu");
  }

  menuToggle.addEventListener("click", () => {
    setMenuOpen(!header.classList.contains("is-menu-open"));
  });

  primaryNav.querySelectorAll("a").forEach((link) => {
    link.addEventListener("click", () => setMenuOpen(false));
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      setMenuOpen(false);
    }
  });
}

// Lerp for smooth parallax interpolation
function lerp(a, b, t) {
  return a + (b - a) * t;
}

// Store current (lerped) and target parallax values per element
const parallaxState = new Map();
parallaxItems.forEach((item) => {
  parallaxState.set(item, { current: 0, target: 0 });
});

const revealObserver = new IntersectionObserver(
  (entries) => {
    entries.forEach((entry) => {
      if (entry.isIntersecting) {
        entry.target.classList.add("is-visible");
      }
    });
  },
  {
    threshold: 0.08,
    rootMargin: "0px 0px -4% 0px",
  },
);

revealItems.forEach((item) => revealObserver.observe(item));

let progressCurrent = 0;

function updateScrollEffects() {
  const maxScroll = document.documentElement.scrollHeight - window.innerHeight;
  const scrollRatio = maxScroll > 0 ? window.scrollY / maxScroll : 0;

  progressCurrent = lerp(progressCurrent, scrollRatio, 0.09);
  progress.style.transform = `scaleX(${progressCurrent})`;

  header.classList.toggle("is-scrolled", window.scrollY > 12);

  parallaxItems.forEach((item) => {
    const speed = Number(item.dataset.parallax || 0);
    const rect = item.getBoundingClientRect();
    const midpoint = rect.top + rect.height / 2 - window.innerHeight / 2;
    const targetY = midpoint * speed * -0.18;
    const state = parallaxState.get(item);
    state.target = targetY;
    state.current = lerp(state.current, state.target, 0.07);
    item.style.transform = `translateY(${state.current}px)`;
  });
}

let ticking = false;

function onScroll() {
  if (!ticking) {
    window.requestAnimationFrame(() => {
      updateScrollEffects();
      ticking = false;
    });
    ticking = true;
  }
}

// Smooth animation loop for lerp continuity
function animationLoop() {
  parallaxItems.forEach((item) => {
    const state = parallaxState.get(item);
    state.current = lerp(state.current, state.target, 0.07);
    item.style.transform = `translateY(${state.current}px)`;
  });
  progressCurrent = lerp(progressCurrent, progressCurrent, 0.09);
  requestAnimationFrame(animationLoop);
}

window.addEventListener("scroll", onScroll, { passive: true });
window.addEventListener("resize", updateScrollEffects);
updateScrollEffects();
requestAnimationFrame(animationLoop);
