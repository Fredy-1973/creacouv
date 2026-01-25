import streamlit as st
import math
import re
from PIL import Image, ImageDraw
import io

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="CréaCouv V9.9.6 PRO", page_icon="🎨")

# --- STYLE CSS POUR LE DESIGN ---
st.markdown("""
    <style>
    .main { opacity: 0.95; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# --- DONNÉES ---
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

tabs = st.tabs(["1. Calculateur de dos", "2. Production Pages", "3. Code-Barres", "4. Aide"])

# --- TAB 1 : CALCULATEUR ---
with tabs[0]:
    st.subheader("Calcul de Couverture & Poids")
    
    col1, col2 = st.columns(2)
    with col1:
        type_papier = st.selectbox("Papier :", list(PAPIERS_DATA.keys()))
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
    st.success(f"**Format ouvert net :** {lt} x {hauteur_f} mm")
    st.warning(f"**Format avec fond perdu (+5mm) :** {lt_fp} x {ht_fp} mm")

    st.write("### Exporter vers votre logiciel :")
    
    # Scripts Generators
    col_a, col_b, col_c = st.columns(3)
    
    # Illustrator JSX
    m = 2.834645
    illu_script = (f'var w={lt_fp}*{m}; var h={ht_fp}*{m}; var doc = app.documents.add(DocumentColorSpace.CMYK, w, h); '
                   f'try {{ doc.views[0].rulerUnits = RulerUnits.Millimeters; }} catch(e) {{}} '
                   f'function dr(x1, y1, x2, y2) {{ var p = doc.pathItems.add(); p.setEntirePath([[x1*{m}, (h-(y1*{m}))], [x2*{m}, (h-(y2*{m}))]]); p.guides = true; }} '
                   f'dr(0, 5, {lt_fp}, 5); dr(0, {ht_fp-5}, {lt_fp}, {ht_fp-5}); dr(5, 0, 5, {ht_fp}); dr({lt_fp-5}, 0, {lt_fp-5}, {ht_fp}); '
                   f'dr({largeur_f+5}, 0, {largeur_f+5}, {ht_fp}); dr({largeur_f+5+dos}, 0, {largeur_f+5+dos}, {ht_fp});')
    
    col_a.download_button("Illustrator (.jsx)", illu_script, file_name="couv_illustrator.jsx")
    
    # InDesign JSX
    indy_script = (f'var d=app.documents.add(); d.documentPreferences.pageHeight="{hauteur_f}mm"; d.documentPreferences.pageWidth="{lt}mm"; '
                   f'd.documentPreferences.documentBleedTopOffset="5mm"; d.documentPreferences.documentBleedUniformSize=true; '
                   f'var m=d.pages[0].marginPreferences; m.top=0; m.bottom=0; m.left=0; m.right=0; '
                   f'd.pages[0].marginPreferences.columnCount=2; d.pages[0].marginPreferences.columnGutter="{dos}mm";')
    col_b.download_button("InDesign (.jsx)", indy_script, file_name="couv_indesign.jsx")

    # PNG Gen (Pillow)
    px = 11.811
    w_px, h_px = int(lt_fp * px), int(ht_fp * px)
    img = Image.new("RGBA", (w_px, h_px), (255, 255, 255, 255))
    draw = ImageDraw.Draw(img)
    draw.rectangle([0, 0, w_px, int(5*px)], fill=(100, 150, 255, 100))
    draw.rectangle([0, h_px-int(5*px), w_px, h_px], fill=(100, 150, 255, 100))
    x_dos = int((largeur_f+5)*px)
    draw.rectangle([x_dos, 0, x_dos + int(dos*px), h_px], fill=(255, 150, 150, 100))
    
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    col_c.download_button("Gabarit Canva (.png)", buf.getvalue(), file_name="gabarit_couv.png")

# --- TAB 2 : PRODUCTION ---
with tabs[1]:
    st.subheader("Présentation Interposeur")
    raw_input = st.text_area("Saisir les numéros de pages couleur (ex: 1, 5, 10-12)", height=100)
    
    if st.button("Calculer binômes"):
        clean = re.sub(r'[^0-9]+', ',', raw_input).strip(',')
        if clean:
            pq = sorted(list(set([int(p) for p in clean.split(',') if p])))
            bn = sorted(list(set([(p if p%2!=0 else p-1, p+1 if p%2!=0 else p) for p in pq])))
            pn = [i + 1 for i, phys in enumerate([p for b in bn for p in b]) if phys not in pq]
            
            res = (f"PAGES COULEUR : {len(pq)}\n"
                   f"FEUILLES COULEUR : {','.join([f'{b[0]}-{b[1]}' for b in bn])}\n"
                   f"POSITIONS NOIRES : {','.join(map(str, sorted(pn)))}")
            st.code(res)
            st.download_button("Exporter Rapport .txt", res, file_name="rapport_prod.txt")

# --- TAB 3 : EAN ---
with tabs[2]:
    st.subheader("Code-barres EAN-13")
    isbn_in = st.text_input("Saisir les 12 premiers chiffres de l'ISBN :", max_chars=12)
    
    if len(isbn_in) == 12:
        s = sum(int(x) * (1 if i % 2 == 0 else 3) for i, x in enumerate(isbn_in))
        full_ean = isbn_in + str((10 - (s % 10)) % 10)
        
        # Logique EAN to PAO
        t = ["AAAAAA", "AABABB", "AABBAB", "AABBBA", "ABAABB", "ABBAAB", "ABBBAA", "ABABAB", "ABABBA", "ABBABA"]
        f = int(full_ean[0]); p = t[f]
        l_part = "".join([chr(65 + int(full_ean[i])) if p[i-1] == 'A' else chr(75 + int(full_ean[i])) for i in range(1, 7)])
        r_part = "".join([chr(97 + int(full_ean[i])) for i in range(7, 13)])
        pao_str = f"{f}{l_part}*{r_part}+"
        
        st.info(f"Code complet : **{full_ean}**")
        st.code(pao_str)
        st.write("Copiez la chaîne ci-dessus et appliquez la police 'Code EAN13'.")

# --- TAB 4 : AIDE ---
with tabs[3]:
    st.info("Support technique : Frédéric Barbier - fred.barbier-icn@outlook.fr")
    st.text("""Utilisation des scripts :
1. Téléchargez le fichier .jsx
2. Dans InDesign/Illustrator/Photoshop, utilisez le menu Fichier > Scripts
3. Le document se crée automatiquement avec les bonnes dimensions.""")