import zipfile
import os
import shutil
import tempfile
import hashlib
import xml.etree.ElementTree as ET
import streamlit as st
from glob import glob

HASH_CIBLE = "591accd6ecdb20315c1ce0017f70029388994ee11bc6fba05a1a53441c6c0240".lower()

def supprimer_image_hash(media_dir):
    supprimés = 0
    for fichier in glob(os.path.join(media_dir, "*.png")):
        try:
            with open(fichier, "rb") as f:
                contenu = f.read()
            sha256 = hashlib.sha256(contenu).hexdigest()
            if sha256 == HASH_CIBLE:
                del contenu
                os.remove(fichier)
                supprimés += 1
        except PermissionError:
            try:
                os.remove(fichier)
                supprimés += 1
            except Exception as e:
                print(f"Impossible de supprimer {fichier}: {e}")
    return supprimés

def supprimer_blocs_pic_preencoded(layout_dir):
    supprimés = 0
    ns = {
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }
    for fichier in glob(os.path.join(layout_dir, "*.xml")):
        tree = ET.parse(fichier)
        root = tree.getroot()
        modifié = False
        for pic in root.findall(".//p:pic", ns):
            cNvPr = pic.find(".//p:nvPicPr/p:cNvPr", ns)
            if cNvPr is not None and cNvPr.get("descr") == "preencoded.png":
                parent = next((elem for elem in root.iter() if pic in list(elem)), None)
                if parent is not None:
                    parent.remove(pic)
                    supprimés += 1
                    modifié = True
        if modifié:
            tree.write(fichier, encoding='utf-8', xml_declaration=True)
    return supprimés

def traiter_pptx(fichier_entree, fichier_sortie, progress_callback=None):
    base_temp = os.path.splitext(os.path.basename(fichier_entree))[0]
    temp_dir = os.path.join(tempfile.gettempdir(), f"{base_temp}_temp_process")
    try:
        if progress_callback:
            progress_callback(1, "Décompression en cours...")
        with zipfile.ZipFile(fichier_entree, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        total_suppr = 0
        layout_dir = os.path.join(temp_dir, 'ppt', 'slideLayouts')
        if os.path.exists(layout_dir):
            if progress_callback:
                progress_callback(2, "Suppression de blocs XML...")
            total_suppr += supprimer_blocs_pic_preencoded(layout_dir)
        media_dir = os.path.join(temp_dir, 'ppt', 'media')
        if os.path.exists(media_dir):
            if progress_callback:
                progress_callback(3, "Suppression des images ciblées...")
            total_suppr += supprimer_image_hash(media_dir)
        if progress_callback:
            progress_callback(4, "Recompression du fichier...")
        with zipfile.ZipFile(fichier_sortie, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    full = os.path.join(root, file)
                    rel = os.path.relpath(full, temp_dir)
                    zip_out.write(full, rel)
        shutil.rmtree(temp_dir)
        return total_suppr
    except Exception as e:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        raise e

st.set_page_config(page_title="Nettoyeur de filigrane Gamma", page_icon="🧹")
st.title("🧹 Nettoyeur de filigrane Gamma (.pptx)")

fichier_entree = st.file_uploader("Choisissez un fichier PPTX d'entrée", type="pptx")
fichier_sortie_nom = st.text_input("Nom du fichier PPTX de sortie", "fichier_modifie.pptx")

progress_bar = st.empty()
status = st.empty()

if fichier_entree and fichier_sortie_nom:
    if st.button("Lancer le nettoyage"):
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_file:
                tmp_file.write(fichier_entree.getvalue())
                tmp_file_path = tmp_file.name
            fichier_sortie_path = os.path.join(tempfile.gettempdir(), fichier_sortie_nom)
            def maj_progression(step, message):
                progress_bar.progress(step / 4)
                status.info(message)
            total = traiter_pptx(tmp_file_path, fichier_sortie_path, maj_progression)
            with open(fichier_sortie_path, "rb") as f:
                st.download_button(
                    label=f"📥 Télécharger ({total} élément(s) supprimé(s))",
                    data=f.read(),
                    file_name=fichier_sortie_nom,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            os.remove(tmp_file_path)
            os.remove(fichier_sortie_path)
        except Exception as e:
            st.error(f"Une erreur est survenue : {str(e)}")
else:
    st.warning("Veuillez choisir un fichier d'entrée et spécifier un fichier de sortie.")