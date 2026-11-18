(() => {
  "use strict";

  const form = document.querySelector("[data-workshop-vars]");
  if (!form) return;

  const storageKey = "workshop-vars";
  const fields = {
    PROXY_HOST: form.querySelector('[name="PROXY_HOST"]'),
    PROXY_PORT: form.querySelector('[name="PROXY_PORT"]'),
    ARTIFACTORY_URL: form.querySelector('[name="ARTIFACTORY_URL"]')
  };

  let stored = {};
  try {
    stored = JSON.parse(localStorage.getItem(storageKey) || "{}");
  } catch (error) {
    stored = {};
  }
  const params = new URLSearchParams(window.location.search);
  Object.keys(fields).forEach((name) => {
    const fromQuery = params.get(name);
    fields[name].value = fromQuery !== null ? fromQuery : stored[name] || "";
  });

  const clean = (name) => {
    let value = fields[name].value.trim();
    if (name === "ARTIFACTORY_URL") value = value.replace(/\/+$/, "");
    return value;
  };

  const values = () => {
    const url = clean("ARTIFACTORY_URL");
    const host = url.replace(/^[a-z]+:\/\//i, "").split(/[/:]/)[0];
    return {
      PROXY_HOST: clean("PROXY_HOST"),
      PROXY_PORT: clean("PROXY_PORT"),
      ARTIFACTORY_URL: url,
      ARTIFACTORY_HOST: host
    };
  };

  const pattern = /ARTIFACTORY_HOST|ARTIFACTORY_URL|PROXY_HOST|PROXY_PORT/g;
  const originals = new Map();
  const roots = document.querySelectorAll(".main-content pre");
  const nodes = [];
  roots.forEach((root) => {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const node = walker.currentNode;
      if (!originals.has(node) && /PROXY_HOST|PROXY_PORT|ARTIFACTORY_/.test(node.nodeValue)) {
        originals.set(node, node.nodeValue);
        nodes.push(node);
      }
    }
  });

  const render = () => {
    const current = values();
    nodes.forEach((node) => {
      node.nodeValue = originals.get(node).replace(pattern, (token) => current[token] || token);
    });
  };

  form.addEventListener("input", () => {
    const toSave = {};
    Object.keys(fields).forEach((name) => { toSave[name] = fields[name].value; });
    try {
      localStorage.setItem(storageKey, JSON.stringify(toSave));
    } catch (error) {
      // Values still apply for this page view when storage is unavailable.
    }
    render();
  });
  form.addEventListener("submit", (event) => event.preventDefault());

  render();
})();
