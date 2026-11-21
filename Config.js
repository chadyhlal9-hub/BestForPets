/* ============================================================
   RÉGLAGES DU SITE – c’est le seul fichier à modifier
   pour les coordonnées, les magasins et la livraison.
   Les produits, eux, se gèrent dans Best_for_Pets.xlsx
   ============================================================ */
const CONFIG = {
  // Fichier Excel du catalogue et dossier des photos
  fichierExcel: "Best_for_Pets.xlsx",
  dossierPhotos: "photos/",
  devise: "DT",

  // Livraison
  livraisonGratuiteDes: 150, // montant (DT) à partir duquel la livraison est offerte
  fraisLivraison: 8,         // frais de livraison en dessous de ce montant

  // Coordonnées (laissez "" ce que vous n’avez pas : ça s’affiche seulement si rempli)
  whatsapp: "21696319915",   // format international sans + ni espaces, ex : "21620123456"
  telephone: "+216 27 737 442",  // ex : "+216 20 123 456"
  email: "BestForPets@gmail.com",      // ex : "contact@exemple.tn"
  facebook: "https://www.facebook.com/share/1DpQgL8NRM/",   // lien complet de la page
  tiktok: "https://www.tiktok.com/@riahi_hamdi?_r=1&_t=ZS-9BX4i3y3H1t",     // lien complet du compte

  // Vos 2 magasins
  magasins: [
    { nom: "Boutique 1 ", adresse: "Oued Ellil - Manouba", telephone: "+216 27 737 442", horaires: "Lundi-Samedi 8h-17h", maps: "https://maps.app.goo.gl/h8G5JXEpbwbR85re9", photo: "magasin1.jpg" },
    { nom: "Boutique 2 ", adresse: "Moncef Bey - Tunis", telephone: "+216 27 737 442", horaires: "Dimache 8h-12h", maps: "https://maps.app.goo.gl/MMVzVkR7RNG9s2pW8", photo: "magasin2.jpg" }
  ],

  // Corrections automatiques de noms de catégories écrits dans l’Excel
  corrections: { "Nouritures": "Nourritures" }
};