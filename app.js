(() => {
  const $ = (s) => document.querySelector(s);
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const fmt = (n) => n.toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " " + CONFIG.devise;
  const norm = (s) => String(s).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]/g, "");
  const LOGO = "best-for-pets-aquatics-logo.svg";

  let products = [];
  let cart = [];
  let activeCat = "Tous";
  try { cart = JSON.parse(localStorage.getItem("bfp_cart") || "[]"); } catch (e) { cart = []; }

  /* ---------- Catalogue depuis Excel ---------- */
  async function loadProducts() {
    try {
      if (typeof XLSX === "undefined") {
        throw new Error("La bibliothèque de lecture Excel (XLSX) n’a pas été chargée. Vérifiez l’accès à Internet ou au CDN dans index.html.");
      }
      const res = await fetch(CONFIG.fichierExcel + "?v=" + Date.now());
      if (!res.ok) throw new Error("Le fichier " + CONFIG.fichierExcel + " est inaccessible (HTTP " + res.status + "). Vérifiez son nom et son emplacement sur le serveur.");
      const wb = XLSX.read(await res.arrayBuffer(), { type: "array" });
      if (!wb.SheetNames.length || !wb.Sheets[wb.SheetNames[0]]) {
        throw new Error("Le fichier Excel ne contient aucune feuille lisible.");
      }
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
      products = rows.map(toProduct).filter(Boolean);
      if (!products.length) {
        throw new Error("Le fichier Excel a été ouvert, mais aucun produit valide n’a été trouvé. Vérifiez que la première ligne contient les colonnes nom et prix.");
      }
      buildCatalogue();
    } catch (err) {
      console.error(err);
      const reason = err instanceof Error ? err.message : String(err);
      const fileOpenedDirectly = location.protocol === "file:";
      const help = fileOpenedDirectly
        ? "La page est ouverte directement depuis un fichier. Lancez-la avec un serveur local ou hébergez le site, puis rechargez la page."
        : "Ouvrez les outils de développement du navigateur (F12), puis l’onglet Console ou Réseau pour voir le détail.";
      $("#grid").innerHTML = `<p class="state"><b>Impossible de charger les produits.</b><br>${esc(reason)}<br>${esc(help)}</p>`;
    }
  }

  /* Photos : le nom dans l’Excel peut s’écrire avec ou sans extension (.jpg, .png…).
     Le site essaie plusieurs variantes, puis affiche le logo si aucune photo n’est trouvée. */
  const EXT = /\.(jpe?g|png|webp|gif|avif|svg)$/i;
  function photoList(file) {
    const base = String(file || "").trim();
    if (!base) return [];
    const names = EXT.test(base) ? [base] : ["jpg", "jpeg", "png", "webp", "JPG", "JPEG", "PNG"].map((e) => base + "." + e);
    const all = [...new Set(names.concat(names.map((n) => n.toLowerCase())))];
    return all.map((n) => CONFIG.dossierPhotos + n.split("/").map(encodeURIComponent).join("/"));
  }
  document.addEventListener("error", (e) => {
    const im = e.target;
    if (!(im instanceof HTMLImageElement) || !im.classList.contains("photo")) return;
    let rest = [];
    try { rest = JSON.parse(im.dataset.alt || "[]"); } catch (x) {}
    if (rest.length) { im.dataset.alt = JSON.stringify(rest.slice(1)); im.src = rest[0]; }
    else im.outerHTML = `<img class="ph" src="${LOGO}" alt="">`;
  }, true);
  /* Photos de carte des magasins : si le fichier est introuvable, la photo disparaît simplement. */
  document.addEventListener("error", (e) => {
    const im = e.target;
    if (!(im instanceof HTMLImageElement) || !im.closest(".store__map")) return;
    let rest = [];
    try { rest = JSON.parse(im.dataset.alt || "[]"); } catch (x) {}
    if (rest.length) { im.dataset.alt = JSON.stringify(rest.slice(1)); im.src = rest[0]; }
    else im.closest(".store__map").remove();
  }, true);

  function toProduct(row) {
    const o = {};
    for (const k in row) o[norm(k)] = row[k];
    const priceKey = Object.keys(o).find((k) => k.startsWith("prix"));
    const name = String(o.nom ?? "").trim();
    const price = parseFloat(String(o[priceKey] ?? "").replace(",", "."));
    if (!name || isNaN(price)) return null;
    let cat = String(o.categorie ?? "").trim() || "Autres";
    cat = CONFIG.corrections[cat] || cat;
    const s = norm(o.stock || "en stock");
    const status = s.includes("rupture") || s.includes("epuis") || s.includes("horsstock") || s.includes("indispo") || s === "non" ? "out" : s.includes("commande") ? "order" : "in";
    return {
      name, price, cat, status,
      photo: String(o.nomfichierphoto ?? "").trim()
    };
  }

  function buildCatalogue() {
    const cats = ["Tous", ...new Set(products.map((p) => p.cat))];
    $("#chips").innerHTML = cats.map((c) => `<button class="chip${c === activeCat ? " active" : ""}" data-cat="${esc(c)}">${esc(c)}</button>`).join("");
    $("#stats").innerHTML = `
      <div><strong>${products.length}</strong><span>produits</span></div>
      <div><strong>${cats.length - 1}</strong><span>catégories</span></div>
      <div><strong>${CONFIG.magasins.length}</strong><span>magasins</span></div>`;
    renderGrid();
  }

  function renderGrid() {
    const q = norm($("#search").value);
    const sort = $("#sort").value;
    let list = products.filter((p) => (activeCat === "Tous" || p.cat === activeCat) && (!q || norm(p.name + p.cat).includes(q)));
    if (sort === "asc") list.sort((a, b) => a.price - b.price);
    if (sort === "desc") list.sort((a, b) => b.price - a.price);
    if (sort === "az") list.sort((a, b) => a.name.localeCompare(b.name, "fr"));
    if (!list.length) { $("#grid").innerHTML = `<p class="state">Aucun produit trouvé.</p>`; return; }
    const labels = { in: "En stock", order: "Sur commande", out: "Hors stock" };
    $("#grid").innerHTML = list.map((p) => {
      const i = products.indexOf(p);
      const photos = photoList(p.photo);
      const img = photos.length
        ? `<img class="photo" src="${esc(photos[0])}" data-alt="${esc(JSON.stringify(photos.slice(1)))}" alt="${esc(p.name)}" loading="lazy">`
        : `<img class="ph" src="${LOGO}" alt="">`;
      const out = p.status === "out";
      return `<article class="card${out ? " card--out" : ""}">
        <div class="card__img">${img}<span class="tag">${esc(p.cat)}</span><span class="stock stock--${p.status}">${labels[p.status]}</span></div>
        <div class="card__body">
          <h3 class="card__name">${esc(p.name)}</h3>
          <div class="card__foot">
            <span class="price">${fmt(p.price)}</span>
            <button class="add${out ? " add--out" : ""}" data-i="${i}" ${out ? "disabled" : ""}>${out ? "Hors stock" : "Ajouter"}</button>
          </div>
        </div></article>`;
    }).join("");
  }

  /* ---------- Panier ---------- */
  const save = () => { try { localStorage.setItem("bfp_cart", JSON.stringify(cart)); } catch (e) {} };
  const subtotal = () => cart.reduce((s, l) => s + l.price * l.qty, 0);
  const shipping = (mode) => (mode === "retrait" || subtotal() >= CONFIG.livraisonGratuiteDes || !cart.length ? 0 : CONFIG.fraisLivraison);

  function addToCart(p) {
    const line = cart.find((l) => l.name === p.name);
    line ? line.qty++ : cart.push({ name: p.name, price: p.price, qty: 1 });
    save(); renderCart();
    const b = $("#openCart"); b.classList.remove("bump"); void b.offsetWidth; b.classList.add("bump");
    toast("« " + p.name + " » ajouté au panier");
  }

  function renderCart() {
    const count = cart.reduce((s, l) => s + l.qty, 0);
    $("#cartCount").textContent = count;
    const sub = subtotal(), ship = shipping("livraison");
    $("#cartItems").innerHTML = cart.length
      ? cart.map((l, i) => `<div class="line">
          <span class="line__name">${esc(l.name)}</span><span class="line__price">${fmt(l.price * l.qty)}</span>
          <div class="qty"><button data-act="minus" data-i="${i}" aria-label="Retirer un">−</button><span>${l.qty}</span><button data-act="plus" data-i="${i}" aria-label="Ajouter un">+</button></div>
          <button class="rm" data-act="rm" data-i="${i}">Supprimer</button></div>`).join("")
      : `<div class="empty"><img src="${LOGO}" alt=""><b>Votre panier est vide</b>Découvrez nos produits et ajoutez ce qui vous plaît !</div>`;
    $("#subTotal").textContent = fmt(sub);
    $("#shipCost").textContent = !cart.length ? "—" : ship === 0 ? "Offerte" : fmt(ship);
    $("#grandTotal").textContent = fmt(sub + ship);
    $("#checkoutBtn").disabled = !cart.length;
    const left = CONFIG.livraisonGratuiteDes - sub;
    $("#shipText").textContent = left > 0 ? `Plus que ${fmt(left)} pour la livraison offerte` : "Livraison offerte sur cette commande !";
    $("#shipBar").style.width = Math.min(100, (sub / CONFIG.livraisonGratuiteDes) * 100) + "%";
  }

  function openCart(open) {
    $("#drawer").classList.toggle("open", open);
    $("#overlay").classList.toggle("show", open);
    $("#drawer").setAttribute("aria-hidden", String(!open));
  }

  /* ---------- Commande ---------- */
  function openModal(open) {
    $("#modal").classList.toggle("show", open);
    $("#modal").setAttribute("aria-hidden", String(!open));
    if (open) { $("#orderView").hidden = false; $("#doneView").hidden = true; $("#formError").textContent = ""; updateRecap(); }
  }

  function updateRecap() {
    const mode = $("#mode").value;
    const isPickup = mode === "retrait";
    $("#deliveryFields").hidden = isPickup;
    $("#storeField").hidden = !isPickup;
    const sub = subtotal(), ship = shipping(mode);
    $("#recap").innerHTML = cart.map((l) => `<div><span>${l.qty} × ${esc(l.name)}</span><span>${fmt(l.price * l.qty)}</span></div>`).join("")
      + `<div><span>Livraison</span><span>${isPickup ? "Retrait gratuit" : ship === 0 ? "Offerte" : fmt(ship)}</span></div>`
      + `<div class="t"><span>Total à payer</span><span>${fmt(sub + ship)}</span></div>`;
  }

  function submitOrder(e) {
    e.preventDefault();
    const f = new FormData($("#orderForm"));
    const get = (k) => String(f.get(k) || "").trim();
    const mode = get("mode"), err = $("#formError");
    if (get("nom").length < 3) return (err.textContent = "Veuillez saisir votre nom et prénom.");
    if (!/^[0-9+ ]{8,16}$/.test(get("tel"))) return (err.textContent = "Veuillez saisir un numéro de téléphone valide.");
    if (mode === "livraison" && (get("ville").length < 2 || get("adresse").length < 5)) return (err.textContent = "Veuillez saisir votre ville et votre adresse complète.");
    err.textContent = "";

    const sub = subtotal(), ship = shipping(mode);
    const store = CONFIG.magasins[+get("magasin")] || {};
    const lines = [
      "*Nouvelle commande – BEST FOR PETS (aquatics)*", "",
      ...cart.map((l) => `• ${l.qty} × ${l.name} : ${fmt(l.price * l.qty)}`), "",
      `Sous-total : ${fmt(sub)}`,
      `Livraison : ${mode === "retrait" ? "Retrait en magasin" : ship === 0 ? "Offerte" : fmt(ship)}`,
      `*Total : ${fmt(sub + ship)}*`, "",
      `Nom : ${get("nom")}`, `Téléphone : ${get("tel")}`,
      mode === "livraison" ? `Adresse : ${get("adresse")}, ${get("ville")}` : `Retrait au : ${store.nom || "magasin"}`,
      get("note") ? `Remarque : ${get("note")}` : ""
    ].filter((l, i, a) => l !== "" || a[i - 1] !== "");
    window.open("https://wa.me/" + CONFIG.whatsapp.replace(/\D/g, "") + "?text=" + encodeURIComponent(lines.join("\n")), "_blank");

    cart = []; save(); renderCart();
    $("#orderForm").reset();
    $("#orderView").hidden = true; $("#doneView").hidden = false;
    openCart(false);
  }

  /* ---------- Contenu depuis config ---------- */
  function renderStatic() {
    const a = [`Livraison offerte dès ${fmt(CONFIG.livraisonGratuiteDes)}`, "Commande simple via WhatsApp"];
    const message = a.map((t) => `<span class="announce__item">${esc(t)}</span>`).join('<span class="announce__separator" aria-hidden="true">•</span>');
    const group = `<div class="announce__group">${message}</div>`;
    $("#announce").innerHTML = `<div class="announce__track">${group.repeat(4)}</div>`;
    $("#heroBadge").innerHTML = `<div>Livraison offerte dès<b>${CONFIG.livraisonGratuiteDes} ${CONFIG.devise}</b></div>`;
    $("#year").textContent = new Date().getFullYear();

    $("#stores").innerHTML = CONFIG.magasins.map((m) => `<div class="store"><h3>${esc(m.nom)}</h3>
      ${m.adresse ? `<p><b>Adresse :</b> ${esc(m.adresse)}</p>` : ""}
      ${m.telephone ? `<p><b>Téléphone :</b> <a href="tel:${esc(m.telephone.replace(/\s/g, ""))}">${esc(m.telephone)}</a></p>` : ""}
      ${m.horaires ? `<p><b>Horaires :</b> ${esc(m.horaires)}</p>` : ""}
      ${photoList(m.photo).length ? `<a class="store__map" href="${esc(m.maps || "#")}" target="_blank" rel="noopener" aria-label="Voir ${esc(m.nom)} sur la carte"><img src="${esc(photoList(m.photo)[0])}" data-alt="${esc(JSON.stringify(photoList(m.photo).slice(1)))}" alt="Carte : ${esc(m.nom)}" loading="lazy"></a>` : ""}
      ${m.maps ? `<a class="btn btn--primary" href="${esc(m.maps)}" target="_blank" rel="noopener">Itinéraire</a>` : ""}</div>`).join("");

    $("#storeSelect").innerHTML = CONFIG.magasins.map((m, i) => `<option value="${i}">${esc(m.nom)}${m.adresse ? " – " + esc(m.adresse) : ""}</option>`).join("");

    const contact = [
      CONFIG.telephone && `<li><a href="tel:${esc(CONFIG.telephone.replace(/\s/g, ""))}">${esc(CONFIG.telephone)}</a></li>`,
      CONFIG.whatsapp && `<li><a href="https://wa.me/${esc(CONFIG.whatsapp.replace(/\D/g, ""))}" target="_blank" rel="noopener">WhatsApp</a></li>`,
      CONFIG.email && `<li><a href="mailto:${esc(CONFIG.email)}">${esc(CONFIG.email)}</a></li>`
    ].filter(Boolean);
    $("#contactLinks").innerHTML = contact.join("") || "<li>Bientôt disponible</li>";
    const social = [
      CONFIG.facebook && `<li><a href="${esc(CONFIG.facebook)}" target="_blank" rel="noopener">Facebook</a></li>`,
      CONFIG.tiktok && `<li><a href="${esc(CONFIG.tiktok)}" target="_blank" rel="noopener">TikTok</a></li>`
    ].filter(Boolean);
    $("#socialLinks").innerHTML = social.join("") || "<li>Bientôt disponible</li>";
  }

  let tt;
  function toast(msg) {
    const t = $("#toast"); t.textContent = msg; t.classList.add("show");
    clearTimeout(tt); tt = setTimeout(() => t.classList.remove("show"), 2200);
  }

  /* ---------- Événements ---------- */
  document.addEventListener("click", (e) => {
    const t = e.target.closest("button, a");
    if (!t) return;
    if (t.classList.contains("chip")) {
      activeCat = t.dataset.cat;
      document.querySelectorAll(".chip").forEach((c) => c.classList.toggle("active", c === t));
      renderGrid();
    } else if (t.classList.contains("add")) addToCart(products[+t.dataset.i]);
    else if (t.dataset.act) {
      const i = +t.dataset.i;
      if (t.dataset.act === "plus") cart[i].qty++;
      if (t.dataset.act === "minus") cart[i].qty = Math.max(1, cart[i].qty - 1);
      if (t.dataset.act === "rm") cart.splice(i, 1);
      save(); renderCart();
    } else if (t.closest("#nav")) $("#nav").classList.remove("open");
  });
  $("#openCart").onclick = () => openCart(true);
  $("#closeCart").onclick = $("#continueBtn").onclick = $("#overlay").onclick = () => openCart(false);
  $("#checkoutBtn").onclick = () => { openCart(false); openModal(true); };
  $("#closeModal").onclick = $("#doneBtn").onclick = () => openModal(false);
  $("#modal").addEventListener("click", (e) => { if (e.target.id === "modal") openModal(false); });
  $("#mode").onchange = updateRecap;
  $("#orderForm").addEventListener("submit", submitOrder);
  $("#search").oninput = $("#sort").onchange = renderGrid;
  $("#burger").onclick = () => $("#nav").classList.toggle("open");
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") { openCart(false); openModal(false); } });

  renderStatic();
  renderCart();
  loadProducts();
})();
