const observedItems = document.querySelectorAll(
  ".product-frame, .feature-grid article, .statement-card, .step-list li, .method-grid article, .large-product, .support-grid article, .ready-badge, .product-proof, .product-dashboard article"
);

const reveal = (entry) => {
  if (entry.isIntersecting) {
    entry.target.classList.add("is-visible");
  }
};

if ("IntersectionObserver" in window) {
  const observer = new IntersectionObserver(
    (entries) => entries.forEach(reveal),
    { threshold: 0.16 }
  );

  observedItems.forEach((item) => {
    item.classList.add("reveal");
    observer.observe(item);
  });
} else {
  observedItems.forEach((item) => item.classList.add("is-visible"));
}
