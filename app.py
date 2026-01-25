import streamlit as st
import math
import re
from PIL import Image, ImageDraw
import io

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="CréaCouv V9.9.6 PRO", page_icon="🎨")

# --- DONNÉES PAPIERS ---
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

st.title("🎨 CréaCouv V9.9.6 PRO")

tabs = st.tabs(["1. Calculateur de dos", "2. Production Pages", "3. Code-Barres", "4. Aide & Notice"])

# --- TAB 1 : CALCULATEUR & EXPORTS ---
with tabs[0]:
    st.subheader("Calcul de Couverture & Poids")
    
    col1, col2 = st.columns(2)
    with col1:
        type_papier = st.selectbox("Sélectionner le papier :", list(PAPIERS_DATA.keys()))
        largeur_f = st.number_input("Largeur finie (mm) :", value=148.0)
    with col2:
        nb_pages = st.number_input("Nombre de pages :", value=100, step=2)
        hauteur_f = st.number_input("Hauteur finie (mm) :", value=210.0)

    # Logique de calcul
    data = PAPIERS_DATA[type_papier]
    dos = math.ceil(round(nb_pages * data["main"], 2))
    if dos < 3: dos = 3
    lt = (2 * largeur_f) + dos
    lt_fp, ht_fp = lt + 10, hauteur_f + 10

    st.divider()
    st.metric("Dos calculé", f"{dos} mm")
    st.info(f"**Format ouvert net :** {lt} x {hauteur_f} mm")
    st.warning(f"**Format avec fond perdu (+5mm) :** {lt_fp} x {ht_fp} mm")

    st.write("### 📂 Générer les scripts de montage")
    
    c1, c2, c3 = st.columns(3)
    m = 2.834645 # Conversion mm vers points pour scripts Adobe
    
    # InDesign
    indy = (f'var d=app.documents.add(); d.documentPreferences.pageHeight="{hauteur_f}mm"; d.documentPreferences.pageWidth="{lt}mm"; '
            f'd.documentPreferences.documentBleedTopOffset="5mm"; d.documentPreferences.documentBleedUniformSize=true; '
            f'var m=d.pages[0].marginPreferences; m.top=0; m.bottom=0; m.left=0; m.right=0; '
            f'd.pages[0].marginPreferences.columnCount=2; d.pages[0].marginPreferences.columnGutter="{dos}mm";')
    c1.download_button("InDesign (.jsx)", indy, file_name="creacouv_indesign.jsx")

    # Photoshop
    photo = (f'app.preferences.rulerUnits = Units.MM; var doc = app.documents.add({lt}+10, {hauteur_f}+10, 300, "Couv", NewDocumentMode.CMYK); '
             f'function g(p, o){{ doc.guides.add(o, p); }} g(5, Direction.HORIZONTAL); g({hauteur_f}+5, Direction.HORIZONTAL); '
             f'g(5, Direction.VERTICAL); g({lt}+5, Direction.VERTICAL); g({largeur_f}+5, Direction.VERTICAL); g({largeur_f+dos}+5, Direction.VERTICAL);')
    c2.download_button("Photoshop (.jsx)", photo, file_name="creacouv_photoshop.jsx")

    # Illustrator
    illu = (f'var w={lt_fp}*{m}; var h={ht_fp}*{m}; var doc = app.documents.add(DocumentColorSpace.CMYK, w, h); '
            f'try {{ doc.views[0].rulerUnits = RulerUnits.Millimeters; }} catch(e) {{}} '
            f'function dr(x1, y1, x2, y2) {{ var p = doc.pathItems.add(); p.setEntirePath([[x1*{m}, (h-(y1*{m}))], [x2*{m}, (h-(y2*{m}))]]); p.guides = true; }} '
            f'dr(0, 5, {lt_fp}, 5); dr(0, {ht_fp-5}, {lt_fp}, {ht_fp-5}); dr(5, 0, 5, {ht_fp}); dr({lt_fp-5}, 0, {lt_fp-5}, {ht_fp}); '
            f'dr({largeur_f+5}, 0, {largeur_f+5}, {ht_fp}); dr({largeur_f+5+dos}, 0, {largeur_f+5+dos}, {ht_fp});')
    c3.download_button("Illustrator (.jsx)", illu, file_name="creacouv_illustrator.jsx")

    c4, c5, c6 = st.columns(3)
    
    # Word (VBS) - Note: L'utilisateur devra l'exécuter sur Windows
    word_vbs = (f'Set w=CreateObject("Word.Application"): w.Visible=True: Set d=w.Documents.Add: With d.PageSetup: .PageWidth={(lt+10)*2.8346}: .PageHeight={(hauteur_f+10)*2.8346}: .TopMargin=0: .BottomMargin=0: .LeftMargin=0: .RightMargin=0: End With: '
                f'Sub L(x1,y1,x2,y2): d.Shapes.AddLine x1,y1,x2,y2: End Sub: L 0,14.17,{(lt+10)*2.8346},14.17: L 0,{(hauteur_f+5)*2.8346},{(lt+10)*2.8346},{(hauteur_f+5)*2.8346}: L 14.17,0,14.17,{(hauteur_f+10)*2.8346}: L {(lt+5)*2.8346},0,{(lt+5)*2.8346},{(hauteur_f+10)*2.8346}: '
                f'L {(largeur_f+5)*2.8346},0,{(largeur_f+5)*2.8346},{(hauteur_f+10)*2.8346}: L {(largeur_f+5+dos)*2.8346},0,{(largeur_f+5+dos)*2.8346},{(hauteur_f+10)*2.8346}')
    c4.download_button("Word (.vbs)", word_vbs, file_name="creacouv_word.vbs")

    # Scribus
    scribus = (f'import scribus\nscribus.newDocument(({lt+10}, {hauteur_f+10}), (0,0,0,0), 0, 1, 1, 0, 0, 1)\n'
               f'scribus.setVGuides([5, {largeur_f+5}, {largeur_f+5+dos}, {lt+5}])\n'
               f'scribus.setHGuides([5, {hauteur_f+5}])')
    c5.download_button("Scribus (.py)", scribus, file_name="creacouv_scribus.py")

    # PNG pour Canva/Affinity
    px_scale = 11.811
    w_px, h_px = int(lt_fp * px_scale), int(ht_fp * px_scale)
    img = Image.new("RGBA", (w_px, h_px), (255, 255, 255, 255))
    draw = ImageDraw.Draw(img)
    draw.rectangle([0, 0, w_px, int(5*px_scale)], fill=(100, 150, 255, 100))
    draw.rectangle([0, h_px-int(5*px_scale), w_px, h_px], fill=(100, 150, 255, 100))
    x_dos = int((largeur_f+5)*px_scale)
    draw.rectangle([x_dos, 0, x_dos + int(dos*px_scale), h_px], fill=(255, 150, 150, 100))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    c6.download_button("Canva/PNG", buf.getvalue(), file_name="gabarit_creacouv.png")

# --- TAB 2 : PRODUCTION ---
with tabs[1]:
    st.subheader("Présentation Interposeur (Quadri/Noir)")
    ent_pages = st.text_area("Saisir la liste des pages couleur physiques :", height=100)
    
    if st.button("Calculer binômes"):
        clean = re.sub(r'[^0-9]+', ',', ent_pages).strip(',')
        if clean:
            pq = sorted(list(set([int(p) for p in clean.split(',') if p])))
            nb_effectif = len(pq)
            bn = sorted(list(set([(p if p%2!=0 else p-1, p+1 if p%2!=0 else p) for p in pq])))
            pn = [i + 1 for i, phys in enumerate([p for b in bn for p in b]) if phys not in pq]
            
            res = (f"NUMÉROS PHYSIQUES RECONNUS :\n{','.join(map(str, pq))}\n"
                   f"NOMBRE DE PAGES COULEUR EFFECTIVES : {nb_effectif}\n"
                   f"--------------------------------------------------\n\n"
                   f"FEUILLES COULEUR (P. Q° / VP) :\n{','.join([f'{b[0]}-{b[1]}' for b in bn])}\n\n"
                   f"POSITIONS NOIRES (P. N° / IP) :\n{','.join(map(str, sorted(pn)))}")
            
            st.text_area("Résultat :", res, height=250)
            st.download_button("Exporter Rapport .txt", res, file_name="rapport_production.txt")

# --- TAB 3 : CODE-BARRES ---
with tabs[2]:
    st.subheader("Générateur EAN-13")
    isbn_12 = st.text_input("Saisir les 12 premiers chiffres de l'ISBN :", max_chars=12)
    
    if len(isbn_12) == 12:
        s = sum(int(x) * (1 if i % 2 == 0 else 3) for i, x in enumerate(isbn_12))
        cksum = (10 - (s % 10)) % 10
        full = isbn_12 + str(cksum)
        
        # EAN to PAO String
        t = ["AAAAAA", "AABABB", "AABBAB", "AABBBA", "ABAABB", "ABBAAB", "ABBBAA", "ABABAB", "ABABBA", "ABBABA"]
        f = int(full[0]); p = t[f]
        l = "".join([chr(65 + int(full[i])) if p[i-1] == 'A' else chr(75 + int(full[i])) for i in range(1, 7)])
        r = "".join([chr(97 + int(full[i])) for i in range(7, 13)])
        pao_final = f"{f}{l}*{r}+"
        
        st.success(f"Code complet : {full}")
        st.info(f"Chaîne PAO : {pao_final}")
        st.write("Appliquez la police **'Code EAN13'** à cette chaîne dans votre logiciel.")

# --- TAB 4 : AIDE COMPLÈTE ---
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
  4.Dans InDesign, double-cliquez sur le script.
  
- Photoshop (.jsx) :
  Dans le logiciel : 'Fichier' > 'Scripts' > 'Parcourir'.

- Illustrator (.jsx) :
  Dans le logiciel : 'Fichier' > 'Scripts' > 'Autre script'.
  Une fois votre maquette créée, allez dans 'Edition', 'Préférences', 'Unités' et sélectionnez 'mm'.

- Word (.vbs) :
  Double-cliquez sur le fichier .vbs. Crée une page au format total avec des lignes de repères pour la coupe et les plis.

- Scribus (.py) :
  'Scripts' > 'Démarrer un script'. Crée le document et ajoute les guides automatiquement.

- Canva / Affinity (.png) :
  Génère une image à 300dpi avec zones de couleurs (Dos et Fond perdu), à placer en fond.

2. PRODUCTION PAGES
-------------------
Saisir la liste des pages couleur physiques, cliquer sur "Calculer binômes" puis sur "Exporter Rapport .txt" et fournir ce fichier texte à votre imprimeur.

3. CODE-BARRES
--------------
Saisir les 12 premiers chiffres de votre N° ISBN (le dernier chiffre est calculé automatiquement).
Copiez la chaîne, collez-la dans votre logiciel de PAO et appliquez la police "Code EAN13" (Corps 30).

Support technique : Frédéric Barbier
Contact : fred.barbier-icn@outlook.fr"""
    
    st.text_area("Notice d'utilisation", help_text, height=600)
