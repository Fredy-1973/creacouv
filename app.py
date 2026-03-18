import sys
import os
import math
import re
from tkinter import messagebox, filedialog, Canvas
from PIL import Image, ImageDraw

# --- DIAGNOSTIC & COMPATIBILITÉ ---
try:
    import customtkinter as ctk
except Exception as e:
    print(f"Erreur : Le module 'custom customtkinter' est manquant. Installez-le avec 'pip install customtkinter'.")
    sys.exit()

# --- CONFIGURATION ---
ctk.set_appearance_mode("light") 
ctk.set_default_color_theme("blue")

def resource_path(relative_path):
    """ Gestion des chemins pour l'export en .EXE (PyInstaller) """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def register_font_if_present(ttf_filename):
    path = resource_path(ttf_filename)
    if not os.path.exists(path): return False
    if sys.platform.startswith("win"):
        try:
            from ctypes import windll
            windll.gdi32.AddFontResourceExW(path, 0x10, 0)
            return True
        except: return False
    return False

class CreacouvApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("CréaCouv V11.3 PRO")
        self.geometry("780x900")
        
        # --- CHARGEMENT DU LOGO (BARRE DE TITRE) ---
        try:
            self.iconbitmap(resource_path("logo.ico"))
        except:
            pass # Si le logo est absent, l'appli se lance normalement

        register_font_if_present("EAN13.TTF")
        
        # Variables
        self.largeur_finie = ctk.StringVar(value="148")
        self.hauteur_finie = ctk.StringVar(value="210")
        self.rabat_val = ctk.StringVar(value="0")
        self.nb_pages = ctk.StringVar(value="100")
        self.type_papier = ctk.StringVar(value="Bouffant 80g")
        self.isbn_var = ctk.StringVar()
        self.pao_str_var = ctk.StringVar()
        
        self.papiers_data = {
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
        
        self.tabview = ctk.CTkTabview(self, width=750, height=850)
        self.tabview.pack(pady=5)
        
        self.tab_cov = self.tabview.add("1. Calculateur de dos")
        self.tab_prod = self.tabview.add("2. Production Pages")
        self.tab_ean = self.tabview.add("3. Code-Barres")
        self.tab_help = self.tabview.add("4. Aide & Notice")
        
        self.setup_cov_tab()
        self.setup_prod_tab()
        self.setup_ean_tab()
        self.setup_help_tab()
        self.bind('<Return>', lambda event: self.calculate_all())

    def setup_cov_tab(self):
        ctk.CTkLabel(self.tab_cov, text="Calcul de Couverture (Rabats inclus)", font=("Arial", 18, "bold")).pack(pady=10)
        f1 = ctk.CTkFrame(self.tab_cov, fg_color="transparent")
        f1.pack(pady=5)
        ctk.CTkLabel(f1, text="Papier :").grid(row=0, column=0, padx=5)
        self.opt_pap = ctk.CTkOptionMenu(f1, values=list(self.papiers_data.keys()), variable=self.type_papier, width=220)
        self.opt_pap.grid(row=0, column=1, padx=10)
        ctk.CTkLabel(f1, text="Pages :").grid(row=0, column=2, padx=5)
        ctk.CTkEntry(f1, textvariable=self.nb_pages, width=80).grid(row=0, column=3, padx=5)

        f2 = ctk.CTkFrame(self.tab_cov, fg_color="transparent")
        f2.pack(pady=5)
        ctk.CTkLabel(f2, text="Largeur :").grid(row=0, column=0, padx=5)
        ctk.CTkEntry(f2, textvariable=self.largeur_finie, width=60).grid(row=0, column=1, padx=5)
        ctk.CTkLabel(f2, text="Hauteur :").grid(row=0, column=2, padx=5)
        ctk.CTkEntry(f2, textvariable=self.hauteur_finie, width=60).grid(row=0, column=3, padx=5)
        ctk.CTkLabel(f2, text="Rabats (mm) :").grid(row=0, column=4, padx=5)
        ctk.CTkEntry(f2, textvariable=self.rabat_val, width=60).grid(row=0, column=5, padx=5)

        ctk.CTkButton(self.tab_cov, text="Calculer", command=self.calculate_all, fg_color="#1F6AA5", font=("Arial", 13, "bold")).pack(pady=15)
        self.lbl_dos = ctk.CTkLabel(self.tab_cov, text="Dos : -- mm", font=("Arial", 18, "bold"), text_color="#1F6AA5"); self.lbl_dos.pack()
        self.lbl_net = ctk.CTkLabel(self.tab_cov, text="Format ouvert net : -- x -- mm", text_color="#E67E22"); self.lbl_net.pack()
        self.lbl_fp = ctk.CTkLabel(self.tab_cov, text="Format avec fond perdu (+5mm) : -- x -- mm"); self.lbl_fp.pack()
        self.lbl_poids = ctk.CTkLabel(self.tab_cov, text="Poids estimé : -- g", font=("Arial", 13, "italic"), text_color="green"); self.lbl_poids.pack(pady=5)

        f_btns = ctk.CTkFrame(self.tab_cov, fg_color="transparent"); f_btns.pack(pady=10)
        self.script_btns = []
        specs = [("InDesign (.jsx)", "#FF3366", self.gen_indesign), ("Photoshop (.jsx)", "#31A8FF", self.gen_photoshop), 
                 ("Illustrator (.jsx)", "#FF7C00", self.gen_illustrator), ("Word (.vbs)", "#2B579A", self.gen_word), 
                 ("Scribus (.py)", "#179C93", self.gen_scribus), ("Canva/PNG", "#6B3FA0", self.gen_affinity)]
        for i, (name, color, cmd) in enumerate(specs):
            b = ctk.CTkButton(f_btns, text=name, command=cmd, state="disabled", fg_color=color, width=160)
            b.grid(row=i//3, column=i%3, padx=5, pady=5); self.script_btns.append(b)

    def calculate_all(self):
        try:
            l = float(self.largeur_finie.get().replace(',', '.'))
            h = float(self.hauteur_finie.get().replace(',', '.'))
            nb = int(self.nb_pages.get())
            r = float(self.rabat_val.get().replace(',', '.'))
            if r > 0 and (r < 60 or r > 120):
                messagebox.showwarning("Rabats", "Les rabats doivent être compris entre 60 et 120 mm.")
                return
            data = self.papiers_data[self.type_papier.get()]
            self.d = max(3, math.ceil(round(nb * data["main"], 2)))
            self.lt = (2 * l) + self.d + (2 * r)
            self.r_val = r
            poids = ((l/1000 * h/1000) * data["gr"] * (nb/2)) + 50
            self.lbl_dos.configure(text=f"Dos : {self.d} mm")
            self.lbl_net.configure(text=f"Format ouvert net : {self.lt} x {h} mm")
            self.lbl_fp.configure(text=f"Format avec fond perdu (+5mm) : {self.lt+10} x {h+10} mm")
            self.lbl_poids.configure(text=f"Poids estimé : {int(poids)} g")
            for b in self.script_btns: b.configure(state="normal")
        except: messagebox.showerror("Erreur", "Saisie invalide.")

    def gen_indesign(self):
        f = filedialog.asksaveasfilename(defaultextension=".jsx")
        if f:
            h, l, r = float(self.hauteur_finie.get().replace(',', '.')), float(self.largeur_finie.get().replace(',', '.')), self.r_val
            s = (f'var d=app.documents.add(); d.documentPreferences.pageHeight="{h}mm"; d.documentPreferences.pageWidth="{self.lt}mm"; '
                 f'd.documentPreferences.documentBleedTopOffset="5mm"; d.documentPreferences.documentBleedUniformSize=true; '
                 f'var p=d.pages[0]; p.marginPreferences.properties={{top:0, left:0, bottom:0, right:0}}; '
                 f'p.guides.add(undefined, {{orientation:HorizontalOrVertical.vertical, location:"{r}mm"}}); '
                 f'p.guides.add(undefined, {{orientation:HorizontalOrVertical.vertical, location:"{r+l}mm"}}); '
                 f'p.guides.add(undefined, {{orientation:HorizontalOrVertical.vertical, location:"{r+l+self.d}mm"}}); '
                 f'p.guides.add(undefined, {{orientation:HorizontalOrVertical.vertical, location:"{r+2*l+self.d}mm"}});')
            with open(f, "w") as file: file.write(s)

    def gen_photoshop(self):
        f = filedialog.asksaveasfilename(defaultextension=".jsx")
        if f:
            l, h, r = float(self.largeur_finie.get().replace(',', '.')), float(self.hauteur_finie.get().replace(',', '.')), self.r_val
            s = (f'app.preferences.rulerUnits = Units.MM; var doc = app.documents.add({self.lt}+10, {h}+10, 300, "Couv", NewDocumentMode.CMYK); '
                 f'function g(p, o){{ doc.guides.add(o, p); }} g(5, Direction.HORIZONTAL); g({h}+5, Direction.HORIZONTAL); '
                 f'g(5, Direction.VERTICAL); g({self.lt}+5, Direction.VERTICAL); g({r}+5, Direction.VERTICAL); '
                 f'g({r+l}+5, Direction.VERTICAL); g({r+l+self.d}+5, Direction.VERTICAL); g({r+2*l+self.d}+5, Direction.VERTICAL);')
            with open(f, "w") as file: file.write(s)

    def gen_illustrator(self):
        f = filedialog.asksaveasfilename(defaultextension=".jsx")
        if f:
            l, h, r = float(self.largeur_finie.get().replace(',', '.')), float(self.hauteur_finie.get().replace(',', '.')), self.r_val
            m = 2.834645
            s = (f'var w=({self.lt}+10)*{m}; var h=({h}+10)*{m}; var doc = app.documents.add(DocumentColorSpace.CMYK, w, h); '
                 f'function dr(x, o){{ var p=doc.pathItems.add(); if(o=="v"){{p.setEntirePath([[x*{m},0],[x*{m},h]]);}}else{{p.setEntirePath([[0,x*{m}],[w,x*{m}]]);}} p.guides=true; }} '
                 f'dr(5, "v"); dr({r+5}, "v"); dr({r+l+5}, "v"); dr({r+l+self.d+5}, "v"); dr({r+2*l+self.d+5}, "v"); dr({self.lt}+5, "v"); '
                 f'dr(5, "h"); dr({h}+5, "h");')
            with open(f, "w") as file: file.write(s)

    def gen_word(self):
        f = filedialog.asksaveasfilename(defaultextension=".vbs")
        if f:
            l, h, r = float(self.largeur_finie.get().replace(',', '.')), float(self.hauteur_finie.get().replace(',', '.')), self.r_val
            lt, ht = self.lt+10, h+10
            s = (f'Set w=CreateObject("Word.Application"): w.Visible=True: Set d=w.Documents.Add: With d.PageSetup: .PageWidth={lt*2.8346}: .PageHeight={ht*2.8346}: .TopMargin=0: .BottomMargin=0: .LeftMargin=0: .RightMargin=0: End With: '
                 f'Sub LV(x): d.Shapes.AddLine x*2.8346, 0, x*2.8346, {ht*2.8346}: End Sub: '
                 f'Sub LH(y): d.Shapes.AddLine 0, y*2.8346, {lt*2.8346}, y*2.8346: End Sub: '
                 f'LV 5: LV {r+5}: LV {r+l+5}: LV {r+l+self.d+5}: LV {r+2*l+self.d+5}: LV {lt-5}: '
                 f'LH 5: LH {h+5}')
            with open(f, "w", encoding="utf-16") as file: file.write(s)

    def gen_scribus(self):
        f = filedialog.asksaveasfilename(defaultextension=".py")
        if f:
            l, h, r = float(self.largeur_finie.get().replace(',', '.')), float(self.hauteur_finie.get().replace(',', '.')), self.r_val
            s = (f'import scribus\nscribus.newDocument(({self.lt+10}, {h+10}), (0,0,0,0), 0, 1, 1, 0, 0, 1)\n'
                 f'scribus.setVGuides([5, {r+5}, {r+l+5}, {r+l+self.d+5}, {r+2*l+self.d+5}, {self.lt+5}])\n'
                 f'scribus.setHGuides([5, {h+5}])')
            with open(f, "w") as file: file.write(s)

    def gen_affinity(self):
        f = filedialog.asksaveasfilename(defaultextension=".png")
        if f:
            l, h, r, px = float(self.largeur_finie.get().replace(',', '.')), float(self.hauteur_finie.get().replace(',', '.')), self.r_val, 11.811
            w_px, h_px = int((self.lt+10)*px), int((h+10)*px)
            img = Image.new("RGBA", (w_px, h_px), (255,255,255,255))
            draw = ImageDraw.Draw(img); draw.rectangle([0,0,w_px,int(5*px)], fill=(100,150,255,100)); draw.rectangle([0,h_px-int(5*px),w_px,h_px], fill=(100,150,255,100))
            if r > 0:
                draw.rectangle([int(5*px),0,int((5+r)*px),h_px], fill=(255,230,100,80)); draw.rectangle([w_px-int((5+r)*px),0,w_px-int(5*px),h_px], fill=(255,230,100,80))
            draw.rectangle([int((5+r+l)*px),0,int((5+r+l+self.d)*px),h_px], fill=(255,150,150,100))
            img.save(f, dpi=(300, 300)); messagebox.showinfo("OK", "PNG 300 DPI généré.")

    def setup_prod_tab(self):
        ctk.CTkLabel(self.tab_prod, text="Production Pages", font=("Arial", 18, "bold")).pack(pady=10)
        self.ent_pages = ctk.CTkTextbox(self.tab_prod, width=550, height=80); self.ent_pages.pack(pady=10)
        f_btns = ctk.CTkFrame(self.tab_prod, fg_color="transparent"); f_btns.pack(pady=10)
        ctk.CTkButton(f_btns, text="Calculer binômes", command=self.calc_interposer, fg_color="#2ECC71").grid(row=0, column=0, padx=5)
        ctk.CTkButton(f_btns, text="Vider", command=self.clear_prod, fg_color="#E74C3C").grid(row=0, column=1, padx=5)
        ctk.CTkButton(f_btns, text="Exporter Rapport .txt", command=self.export_prod, fg_color="gray").grid(row=0, column=2, padx=5)
        self.txt_prod = ctk.CTkTextbox(self.tab_prod, width=650, height=300, font=("Courier", 12)); self.txt_prod.pack(pady=10)

    def clear_prod(self):
        self.ent_pages.delete("0.0", "end")
        self.txt_prod.delete("0.0", "end")

    def calc_interposer(self):
        raw = self.ent_pages.get("0.0", "end")
        clean = re.sub(r'[^0-9]+', ',', raw).strip(',')
        if not clean: return
        pq = sorted(list(set([int(p) for p in clean.split(',') if p])))
        bn = sorted(list(set([(p if p%2!=0 else p-1, p+1 if p%2!=0 else p) for p in pq])))
        pn = [i + 1 for i, phys in enumerate([p for b in bn for p in b]) if phys not in pq]
        res = (f"FEUILLES COULEUR (P. Q° / VP) :\n{','.join([f'{b[0]}-{b[1]}' for b in bn])}\n\n"
               f"POSITIONS NOIRES (P. N° / IP) :\n{','.join(map(str, sorted(pn)))}")
        self.txt_prod.delete("0.0", "end"); self.txt_prod.insert("0.0", res)

    def export_prod(self):
        f = filedialog.asksaveasfilename(defaultextension=".txt")
        if f:
            with open(f, "w") as file: file.write(self.txt_prod.get("0.0", "end"))

    def setup_ean_tab(self):
        ctk.CTkLabel(self.tab_ean, text="Code-barres EAN-13", font=("Arial", 18, "bold")).pack(pady=10)
        self.ent_isbn = ctk.CTkEntry(self.tab_ean, textvariable=self.isbn_var, width=300); self.ent_isbn.pack(pady=10)
        ctk.CTkButton(self.tab_ean, text="Générer", command=self.generate_ean).pack(pady=5)
        self.ent_pao_res = ctk.CTkEntry(self.tab_ean, textvariable=self.pao_str_var, width=400, state="readonly", justify="center"); self.ent_pao_res.pack(pady=20)
        self.canvas_ean = Canvas(self.tab_ean, width=400, height=150, bg="white"); self.canvas_ean.pack(pady=10)

    def generate_ean(self):
        raw = self.isbn_var.get().replace("-", "").strip()
        if len(raw) != 12: return
        s = sum(int(x) * (1 if i % 2 == 0 else 3) for i, x in enumerate(raw))
        full = raw + str((10 - (s % 10)) % 10)
        pao_val = self.ean_to_pao(full); self.pao_str_var.set(pao_val)
        self.canvas_ean.delete("all"); self.canvas_ean.create_text(200, 75, text=pao_val, font=("Code EAN13", 42))

    def ean_to_pao(self, ean):
        t = ["AAAAAA", "AABABB", "AABBAB", "AABBBA", "ABAABB", "ABBAAB", "ABBBAA", "ABABAB", "ABABBA", "ABBABA"]
        f, p = int(ean[0]), t[int(ean[0])]
        l = "".join([chr(65 + int(ean[i])) if p[i-1] == 'A' else chr(75 + int(ean[i])) for i in range(1, 7)])
        r = "".join([chr(97 + int(ean[i])) for i in range(7, 13)])
        return f"{f}{l}*{r}+"

    def setup_help_tab(self):
        txt = ctk.CTkTextbox(self.tab_help, width=700, height=750, font=("Courier", 11)); txt.pack(pady=10)
        help_text = """1. CALCULATEUR DE DOS ET FORMATS
------------------------------
- Dos : Calculé selon le coefficient papier.
- Format Net : Sans fonds perdus.
- Format avec fond perdu : Format total.

UTILISATION DES SCRIPTS :
- InDesign (.jsx) :
  1.'Fenêtre' > 'Utilitaires' > 'Scripts' > 'Utilisateur'.
  2.Panneau 'Scripts' > Clic droit dossier 'Utilisateur' >
  'Faire apparaître dans l'explorateur'.
  3.Glissez le fichier dans le dossier 'Scripts Panel'.
  4.Dans InDesign, fenêtre scripts, double-cliquez sur le script que vous avez créé.
  
- Photoshop (.jsx) :
  Dans le logiciel : 'Fichier' > 'Scripts' > 'Parcourir'.

- Illustrator (.jsx) :
  Dans le logiciel : 'Fichier' > 'Scripts' > 'Autre script'.
  Une fois votre maquette créée, allez dans
  'Edition', 'Préférences', 'Unités' et sélectionnez 'mm'.

- Word (.vbs) :
  Double-cliquez sur le fichier .vbs. Crée une page au format total
  avec des lignes de repères pour la coupe et les plis.

- Scribus (.py) :
  'Scripts' > 'Démarrer un script'. Crée le document et
  ajoute les guides verticaux/horizontaux automatiquement.

- Canva / Affinity (.png) :
  Génère une image à 300dpi avec zones de couleurs (Dos et Fond perdu),
  à placer en fond pour créer votre couverture.

2. PRODUCTION PAGES
-------------------
Saisir la liste des pages couleur physiques, cliquer sur "Calculer binômes"
puis sur "Exporter Rapport .txt" et fournir ce fichier texte à votre imprimeur.

3. CODE-BARRES
--------------
Saisir les 12 premiers chiffres de votre N° ISBN
(le dernier chiffre est calculé automatiquement).
Copiez la chaîne, collez-la dans votre logiciel de PAO
et appliquez la police "Code EAN13" (Corps 30).

Support technique : Frédéric Barbier
Contact : fred.barbier-icn@outlook.fr"""
        txt.insert("0.0", help_text); txt.configure(state="disabled")

if __name__ == "__main__":
    app = CreacouvApp()
    app.mainloop()
