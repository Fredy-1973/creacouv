import streamlit as st
import math
import re
from PIL import Image, ImageDraw
import io

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="CréaCouv V11.7 PRO", page_icon="🎨", layout="centered")

# --- DONNÉES PAPIERS (STRICTEMENT IDENTIQUES) ---
PAPIERS_DATA = {
    "Bouffant 80g": {"main": 0.07, "gr": 80},
    "Bouffant 1.5 80g": {"main": 0.061, "gr": 80},
    "Bouffant 90g": {"main": 0.08, "gr": 90},
    "Offset 60g": {"main": 0.04, "gr": 60},
    "Offset 80g": {"main": 0.05, "gr": 80},
    "Offset 90g": {"main": 0.055, "gr": 90},
    "Munken 70g": {"main": 0.06, "gr": 70},
    "Munken 80g": {"main": 0.075, "gr": 80},
    "Munken 90g": {"main": 0.09, "gr": 90},
    "Munken 115": {"main": 0.11, "gr": 115},
    "Olin 70g": {"main": 0.045, "gr": 70},
    "Olin 80g": {"main": 0.0515, "gr": 80},
    "Olin 90g": {"main": 0.06, "gr": 90},
    "Olin 100g": {"main": 0.0666, "gr": 100},
    "Olin 120g": {"main": 0.0781, "gr": 120},
    "Olin 130g": {"main": 0.09, "gr": 130},
    "Olin 150g": {"main": 0.1, "gr": 150},
    "Couché Brillant 90g": {"main": 0.0333, "gr": 90},
    "Couché Brillant 115g": {"main": 0.0471, "gr": 115},
    "Couché Brillant 135g": {"main": 0.0531, "gr": 135},
    "Couché Brillant 170g": {"main": 0.067, "gr": 170},
    "Couché Mat 90g": {"main": 0.039, "gr": 90},
    "Couché Mat 115g": {"main": 0.052, "gr": 115},
    "Couché Mat 135g": {"main": 0.0625, "gr": 135},
    "Couché Mat 150g": {"main": 0.069, "gr": 150},
    "Couché Mat 170g": {"main": 0.0833, "gr": 170},
    "Couché Mat 200g": {"main": 0.095, "gr": 200}
}

st.title("🎨 CréaCouv V11.7 PRO")
tabs = st.tabs(["1. Calculateur de dos", "2. Production Pages", "3. Code-Barres", "4. Aide & Notice"])

# Initialisation de l'état pour les pages couleur (utile pour le devis)
if 'nb_coul' not in st.session_state:
    st.session_state.nb_coul = 0

# --- TAB 1 : CALCULATEUR ---
with tabs[0]:
    st.subheader("Calcul de Couverture (Rabats inclus)")
    
    col_p1, col_p2 = st.columns([2, 1])
    with col_p1:
        type_papier = st.selectbox("Papier :", list(PAPIERS_DATA.keys()))
    with col_p2:
        nb_pages = st.number_input("Pages :", value=100, step=2)

    col_d1, col_d2, col_d3 = st.columns(3)
    with col_d1:
        largeur_f = st.number_input("Largeur (mm) :", value=148.0)
    with col_d2:
        hauteur_f = st.number_input("Hauteur (mm) :", value=210.0)
    with col_d3:
        rabat_val = st.number_input("Rabats (mm) :", value=0.0)

    # Logique de calcul (Strictement identique à ton script)
    if rabat_val > 0 and (rabat_val < 60 or rabat_val > 120):
        st.error("⚠️ Les rabats doivent être compris entre 60 et 120 mm.")
        dos, lt = 0, 0
    else:
        data = PAPIERS_DATA[type_papier]
        dos = max(3, math.ceil(round(nb_pages * data["main"], 2)))
        lt = (2 * largeur_f) + dos + (2 * rabat_val)
        poids = ((largeur_f/1000 * hauteur_f/1000) * data["gr"] * (nb_pages/2)) + 50

        st.markdown("---")
        # Ajustement des colonnes pour éviter le texte coupé (...)
        c1, c2, c3 = st.columns([25, 50, 25])
        c1.metric("Dos", f"{dos} mm")
        c2.metric("Format ouvert net", f"{lt} x {hauteur_f} mm")
        c3.metric("Poids estimé", f"{int(poids)} g")
        
        st.warning(f"**Format avec fond perdu (+5mm) :** {lt+10} x {hauteur_f+10} mm")
        st.markdown("---")

        # --- EXPORTS ---
        st.write("### 📂 Générer les gabarits")
        m = 2.834645 
        
        e1, e2, e3 = st.columns(3)
        # InDesign
        indy = (f'var d=app.documents.add(); d.documentPreferences.pageHeight="{hauteur_f}mm"; d.documentPreferences.pageWidth="{lt}mm"; '
                f'd.documentPreferences.documentBleedTopOffset="5mm"; d.documentPreferences.documentBleedUniformSize=true; '
                f'var p=d.pages[0]; p.marginPreferences.properties={{top:0, left:"{rabat_val}mm", bottom:0, right:"{rabat_val}mm", columnCount:2, columnGutter:"{dos}mm"}};')
        e1.download_button("InDesign (.jsx)", indy, file_name="creacouv_indesign.jsx")

        # Photoshop
        ps = (f'app.preferences.rulerUnits = Units.MM; var doc = app.documents.add({lt}+10, {hauteur_f}+10, 300, "Couv", NewDocumentMode.CMYK); '
              f'function g(p, o){{ doc.guides.add(o, p); }} g(5, Direction.HORIZONTAL); g({hauteur_f}+5, Direction.HORIZONTAL); '
              f'g(5, Direction.VERTICAL); g({lt}+5, Direction.VERTICAL); g({rabat_val}+5, Direction.VERTICAL); '
              f'g({rabat_val+largeur_f}+5, Direction.VERTICAL); g({rabat_val+largeur_f+dos}+5, Direction.VERTICAL); g({rabat_val+2*largeur_f+dos}+5, Direction.VERTICAL);')
        e2.download_button("Photoshop (.jsx)", ps, file_name="creacouv_photoshop.jsx")

        # Illustrator
        illu = (f'var w=({lt}+10)*{m}; var h=({hauteur_f}+10)*{m}; var doc = app.documents.add(DocumentColorSpace.CMYK, w, h); '
                f'function dr(x, o){{ var p=doc.pathItems.add(); if(o=="v"){{p.setEntirePath([[x*{m},0],[x*{m},h]]);}}else{{p.setEntirePath([[0,x*{m}],[w,x*{m}]]);}} p.guides=true; }} '
                f'dr(5, "v"); dr({rabat_val+5}, "v"); dr({rabat_val+largeur_f+5}, "v"); dr({rabat_val+largeur_f+dos+5}, "v"); dr({rabat_val+2*largeur_f+dos+5}, "v"); dr({lt}+5, "v"); '
                f'dr(5, "h"); dr({hauteur_f}+5, "h");')
        e3.download_button("Illustrator (.jsx)", illu, file_name="creacouv_illustrator.jsx")

        e4, e5, e6 = st.columns(3)
        # Word
        word = (f'Set w=CreateObject("Word.Application"): w.Visible=True: Set d=w.Documents.Add: With d.PageSetup: .PageWidth={(lt+10)*2.8346}: .PageHeight={(hauteur_f+10)*2.8346}: .TopMargin=0: .BottomMargin=0: .LeftMargin=0: .RightMargin=0: End With: '
                f'Sub LV(x): d.Shapes.AddLine x*2.8346, 0, x*2.8346, {(hauteur_f+10)*2.8346}: End Sub: Sub LH(y): d.Shapes.AddLine 0, y*2.8346, {(lt+10)*2.8346}, y*2.8346: End Sub: '
                f'LV 5: LV {rabat_val+5}: LV {rabat_val+largeur_f+5}: LV {rabat_val+largeur_f+dos+5}: LV {rabat_val+2*largeur_f+dos+5}: LV {lt+5}: LH 5: LH {hauteur_f+5}')
        e4.download_button("Word (.vbs)", word, file_name="creacouv_word.vbs")

        # Scribus
        scribus = (f'import scribus\nscribus.newDocument(({lt+10}, {hauteur_f+10}), (0,0,0,0), 0, 1, 1, 0, 0, 1)\n'
                   f'scribus.setVGuides([5, {rabat_val+5}, {rabat_val+largeur_f+5}, {rabat_val+largeur_f+dos+5}, {rabat_val+2*largeur_f+dos+5}, {lt+5}])\nscribus.setHGuides([5, {hauteur_f+5}])')
        e5.download_button("Scribus (.py)", scribus, file_name="creacouv_scribus.py")

        # PNG
        px_val = 11.811
        w_px, h_px = int((lt+10)*px_val), int((hauteur_f+10)*px_val)
        img = Image.new("RGBA", (w_px, h_px), (255,255,255,255))
        draw = ImageDraw.Draw(img)
        draw.rectangle([0,0,w_px,int(5*px_val)], fill=(100,150,255,100))
        draw.rectangle([0,h_px-int(5*px_val),w_px,h_px], fill=(100,150,255,100))
        draw.rectangle([0,0,int(5*px_val),h_px], fill=(100,150,255,100))
        draw.rectangle([w_px-int(5*px_val),0,w_px,h_px], fill=(100,150,255,100))
        if rabat_val > 0:
            draw.rectangle([int(5*px_val),0,int((5+rabat_val)*px_val),h_px], fill=(255,230,100,80))
            draw.rectangle([w_px-int((5+rabat_val)*px_val),0,w_px-int(5*px_val),h_px], fill=(255,230,100,80))
        draw.rectangle([int((5+rabat_val+largeur_f)*px_val),0,int((5+rabat_val+largeur_f+dos)*px_val),h_px], fill=(255,150,150,100))
        buf = io.BytesIO()
        img.save(buf, format="PNG", dpi=(300,300))
        e6.download_button("Canva/PNG", buf.getvalue(), file_name="gabarit_v11_7.png")

        # --- POPUP DEVIS (SIMULÉ) ---
        st.markdown("---")
        if st.button("📄 GÉNÉRER LA DEMANDE DE DEVIS", use_container_width=True):
            nb_noires = int(nb_pages) - st.session_state.nb_coul
            devis_text = f"""Titre du livre : (à compléter)
Nombre de pages noires : {nb_noires}
Nombre de pages couleur : {st.session_state.nb_coul}
Format fini : {largeur_f} x {hauteur_f} mm
Papier intérieur : {type_papier} (à préciser)
Papier couverture : (à renseigner)
Nombre d'exemplaires : (à renseigner par multiple de 4)
Pelliculage : (Mat, Brillant ou Peau de pêche)

--------------------------------------------------
Destinataire : icn@imprimerie-icn.fr
--------------------------------------------------"""
            st.info("**VOTRE RÉCAPITULATIF DEVIS** (Copiez le texte ci-dessous)")
            st.code(devis_text, language="text")

# --- TAB 2 : PRODUCTION ---
with tabs[1]:
    st.subheader("Production Pages")
    raw_prod = st.text_area("Saisir la liste des pages couleur physiques :", height=100)
    
    col_bt1, col_bt2 = st.columns(2)
    if col_bt1.button("Calculer binômes", use_container_width=True):
        clean = re.sub(r'[^0-9]+', ',', raw_prod).strip(',')
        if clean:
            pq = sorted(list(set([int(p) for p in clean.split(',') if p])))
            st.session_state.nb_coul = len(pq)
            bn = sorted(list(set([(p if p%2!=0 else p-1, p+1 if p%2!=0 else p) for p in pq])))
            pn = [i + 1 for i, phys in enumerate([p for b in bn for p in b]) if phys not in pq]
            res = (f"FEUILLES COULEUR (P. Q° / VP) :\n{','.join([f'{b[0]}-{b[1]}' for b in bn])}\n\n"
                   f"POSITIONS NOIRES (P. N° / IP) :\n{','.join(map(str, sorted(pn)))}")
            st.session_state.last_prod = res
            st.success(f"Nombre de pages couleur détectées : {st.session_state.nb_coul}")
            st.code(res)

    if col_bt2.button("Vider", use_container_width=True):
        st.session_state.nb_coul = 0
        st.rerun()

    if 'last_prod' in st.session_state:
        st.download_button("Exporter Rapport .txt", st.session_state.last_prod, file_name="rapport_production.txt")

# --- TAB 3 : EAN ---
with tabs[2]:
    st.subheader("Code-barres EAN-13")
    isbn_in = st.text_input("Saisir les 12 premiers chiffres :", max_chars=12)
    
    if len(isbn_in) == 12:
        s = sum(int(x) * (1 if i % 2 == 0 else 3) for i, x in enumerate(isbn_in))
        full = isbn_in + str((10 - (s % 10)) % 10)
        
        # Logique EAN to PAO
        t_tab = ["AAAAAA", "AABABB", "AABBAB", "AABBBA", "ABAABB", "ABBAAB", "ABBBAA", "ABABAB", "ABABBA", "ABBABA"]
        f_idx, p_str = int(full[0]), t_tab[int(full[0])]
        l_part = "".join([chr(65 + int(full[i])) if p_str[i-1] == 'A' else chr(75 + int(full[i])) for i in range(1, 7)])
        r_part = "".join([chr(97 + int(full[i])) for i in range(7, 13)])
        pao_final = f"{f_idx}{l_part}*{r_part}+"
        
        st.success(f"Code complet : {full}")
        st.info(f"Chaîne PAO : {pao_final}")
        st.caption("Appliquez la police 'Code EAN13' dans votre logiciel.")

        # --- AJOUT DU BOUTON DE TÉLÉCHARGEMENT ---
        st.markdown("---")
        try:
            with open("EAN13.ttf", "rb") as file:
                st.download_button(
                    label="📥 Télécharger la police EAN13",
                    data=file,
                    file_name="EAN13.ttf",
                    mime="font/ttf"
                )
            st.write("💡 **Instruction :** Une fois téléchargé, faites un clic droit sur le fichier et choisissez **'Installer'** pour pouvoir l'utiliser dans votre logiciel de mise en page.")
        except FileNotFoundError:
            st.warning("Note : Le fichier de police EAN13.ttf n'a pas été trouvé sur le serveur.")

# --- TAB 4 : AIDE ---
with tabs[3]:
    help_text = """1. CALCULATEUR DE DOS ET FORMATS
------------------------------
- Dos : Calculé selon le coefficient papier.
- Format Net : Sans fonds perdus.
- Format avec fond perdu : Format total.

UTILISATION DES SCRIPTS :
- InDesign (.jsx) :
  1.'Fenêtre' > 'Utilitaires' > 'Scripts' > 'Utilisateur'.
  2.Panneau 'Scripts' > Clic droit dossier 'Utilisateur' > 'Faire apparaître dans l'explorateur'.
  3.Glissez le fichier dans le dossier 'Scripts Panel'.
  4.Dans InDesign, fenêtre scripts, double-cliquez sur le script que vous avez créé.
  
- Photoshop (.jsx) :
  Dans le logiciel : 'Fichier' > 'Scripts' > 'Parcourir'.

- Illustrator (.jsx) :
  Dans le logiciel : 'Fichier' > 'Scripts' > 'Autre script'.
  Une fois votre maquette créée, allez dans 'Edition', 'Préférences', 'Unités' et sélectionnez 'mm'.

- Word (.vbs) :
  Double-cliquez sur le fichier .vbs. Crée une page au format total avec des lignes de repères pour la coupe et les plis.

- Scribus (.py) :
  'Scripts' > 'Démarrer un script'. Crée le document et ajoute les guides verticaux/horizontaux automatiquement.

- Canva / Affinity (.png) :
  Génère une image à 300dpi avec zones de couleurs (Dos et Fond perdu), à placer en fond pour créer votre couverture.

2. PRODUCTION PAGES
-------------------
Saisir la liste des pages couleur physiques, cliquer sur "Calculer binômes" puis sur "Exporter Rapport .txt" et fournir ce fichier texte à votre imprimeur.

3. CODE-BARRES
--------------
Saisir les 12 premiers chiffres de votre N° ISBN (le dernier chiffre est calculé automatiquement).
Copiez la chaîne, collez-la dans votre logiciel de PAO et appliquez la police "Code EAN13" (Corps 30).

Support technique : Frédéric Barbier
Contact : fred.barbier-icn@outlook.fr"""
    st.text_area("Notice détaillée", help_text, height=500)
