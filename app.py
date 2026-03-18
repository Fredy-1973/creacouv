import streamlit as st
import math
import re
from PIL import Image, ImageDraw
import io

# --- CONFIGURATION ---
st.set_page_config(page_title="CréaCouv V11.3 PRO", page_icon="🎨")

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

st.title("🎨 CréaCouv V11.3 PRO")
tabs = st.tabs(["1. Calculateur de dos", "2. Production Pages", "3. Code-Barres", "4. Aide & Notice"])

# --- TAB 1 : CALCULATEUR ---
with tabs[0]:
    st.subheader("Calcul de Couverture (Rabats inclus)")
    
    col1, col2 = st.columns(2)
    with col1:
        type_papier = st.selectbox("Papier :", list(PAPIERS_DATA.keys()))
        largeur_f = st.number_input("Largeur finie (mm) :", value=148.0)
        rabat_val = st.number_input("Rabats (mm) :", value=0.0)
    with col2:
        nb_pages = st.number_input("Nombre de pages :", value=100, step=2)
        hauteur_f = st.number_input("Hauteur finie (mm) :", value=210.0)

    # Logique de calcul V11.3
    if rabat_val > 0 and (rabat_val < 60 or rabat_val > 120):
        st.error("⚠️ Les rabats doivent être compris entre 60 et 120 mm.")
        dos, lt = 0, 0
    else:
        data = PAPIERS_DATA[type_papier]
        dos = max(3, math.ceil(round(nb_pages * data["main"], 2)))
        lt = (2 * largeur_f) + dos + (2 * rabat_val)
        poids = ((largeur_f/1000 * hauteur_f/1000) * data["gr"] * (nb_pages/2)) + 50

        st.divider()
        c1, c2, c3 = st.columns(3)
        c1.metric("Dos", f"{dos} mm")
        c2.metric("Format Net", f"{lt}x{hauteur_f} mm")
        c3.metric("Poids estimé", f"{int(poids)} g")
        
        st.warning(f"**Format avec fond perdu (+5mm) :** {lt+10} x {hauteur_f+10} mm")

        # --- EXPORTS ---
        st.write("### 📂 Générer les gabarits")
        m = 2.834645
        
        e1, e2, e3 = st.columns(3)
        # InDesign
        indy = (f'var d=app.documents.add(); d.documentPreferences.pageHeight="{hauteur_f}mm"; d.documentPreferences.pageWidth="{lt}mm"; '
                f'd.documentPreferences.documentBleedTopOffset="5mm"; d.documentPreferences.documentBleedUniformSize=true; '
                f'var p=d.pages[0]; p.marginPreferences.properties={{top:0, left:0, bottom:0, right:0}}; '
                f'p.guides.add(undefined, {{orientation:HorizontalOrVertical.vertical, location:"{rabat_val}mm"}}); '
                f'p.guides.add(undefined, {{orientation:HorizontalOrVertical.vertical, location:"{rabat_val+largeur_f}mm"}}); '
                f'p.guides.add(undefined, {{orientation:HorizontalOrVertical.vertical, location:"{rabat_val+largeur_f+dos}mm"}}); '
                f'p.guides.add(undefined, {{orientation:HorizontalOrVertical.vertical, location:"{rabat_val+2*largeur_f+dos}mm"}});')
        e1.download_button("InDesign (.jsx)", indy, file_name="creacouv_v11.jsx")

        # Photoshop
        ps = (f'app.preferences.rulerUnits = Units.MM; var doc = app.documents.add({lt}+10, {hauteur_f}+10, 300, "Couv", NewDocumentMode.CMYK); '
              f'function g(p, o){{ doc.guides.add(o, p); }} g(5, Direction.HORIZONTAL); g({hauteur_f}+5, Direction.HORIZONTAL); '
              f'g(5, Direction.VERTICAL); g({lt}+5, Direction.VERTICAL); g({rabat_val}+5, Direction.VERTICAL); '
              f'g({rabat_val+largeur_f}+5, Direction.VERTICAL); g({rabat_val+largeur_f+dos}+5, Direction.VERTICAL); g({rabat_val+2*largeur_f+dos}+5, Direction.VERTICAL);')
        e2.download_button("Photoshop (.jsx)", ps, file_name="creacouv_v11.jsx")

        # PNG / Canva
        px = 11.811
        w_px, h_px = int((lt+10)*px), int((hauteur_f+10)*px)
        img = Image.new("RGBA", (w_px, h_px), (255,255,255,255))
        draw = ImageDraw.Draw(img)
        draw.rectangle([0,0,w_px,int(5*px)], fill=(100,150,255,100))
        draw.rectangle([0,h_px-int(5*px),w_px,h_px], fill=(100,150,255,100))
        if rabat_val > 0:
            draw.rectangle([int(5*px),0,int((5+rabat_val)*px),h_px], fill=(255,230,100,80))
            draw.rectangle([w_px-int((5+rabat_val)*px),0,w_px-int(5*px),h_px], fill=(255,230,100,80))
        draw.rectangle([int((5+rabat_val+largeur_f)*px),0,int((5+rabat_val+largeur_f+dos)*px),h_px], fill=(255,150,150,100))
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        e3.download_button("Canva/PNG", buf.getvalue(), file_name="gabarit_v11.png")

# --- TAB 2 : PRODUCTION ---
with tabs[1]:
    st.subheader("Production Pages")
    raw_prod = st.text_area("Saisir les pages couleur :", height=100)
    if st.button("Calculer binômes"):
        clean = re.sub(r'[^0-9]+', ',', raw_prod).strip(',')
        if clean:
            pq = sorted(list(set([int(p) for p in clean.split(',') if p])))
            bn = sorted(list(set([(p if p%2!=0 else p-1, p+1 if p%2!=0 else p) for p in pq])))
            pn = [i + 1 for i, phys in enumerate([p for b in bn for p in b]) if phys not in pq]
            res = (f"FEUILLES COULEUR (P. Q° / VP) :\n{','.join([f'{b[0]}-{b[1]}' for b in bn])}\n\n"
                   f"POSITIONS NOIRES (P. N° / IP) :\n{','.join(map(str, sorted(pn)))}")
            st.code(res)

# --- TAB 3 : EAN ---
with tabs[2]:
    st.subheader("Code-barres EAN-13")
    isbn_in = st.text_input("Saisir les 12 premiers chiffres :", max_chars=12)
    if len(isbn_in) == 12:
        s = sum(int(x) * (1 if i % 2 == 0 else 3) for i, x in enumerate(isbn_in))
        full = isbn_in + str((10 - (s % 10)) % 10)
        t = ["AAAAAA", "AABABB", "AABBAB", "AABBBA", "ABAABB", "ABBAAB", "ABBBAA", "ABABAB", "ABABBA", "ABBABA"]
        f, p = int(full[0]), t[int(full[0])]
        l_part = "".join([chr(65 + int(full[i])) if p[i-1] == 'A' else chr(75 + int(full[i])) for i in range(1, 7)])
        r_part = "".join([chr(97 + int(full[i])) for i in range(7, 13)])
        pao = f"{f}{l_part}*{r_part}+"
        st.success(f"Code : {full}")
        st.info(f"Chaîne PAO : {pao}")
        st.caption("Appliquez la police 'Code EAN13' dans votre logiciel.")

# --- TAB 4 : AIDE ---
with tabs[3]:
    st.markdown("""### Notice d'utilisation V11.3
**1. Rabats :** Doivent être entre 60 et 120 mm ou laissés à 0.
**2. Poids :** Estimation incluant 50g de couverture.
**3. Scripts :** Téléchargez et exécutez dans vos logiciels Adobe ou Scribus.
**Support :** Frédéric Barbier (fred.barbier-icn@outlook.fr)""")
