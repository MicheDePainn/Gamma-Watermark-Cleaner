import zipfile
import os
import shutil
import tempfile
import streamlit as st
import xml.etree.ElementTree as ET
from glob import glob

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
        if progress_callback: progress_callback(1, "Décompression en cours...")
        with zipfile.ZipFile(fichier_entree, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        layout_dir = os.path.join(temp_dir, 'ppt', 'slideLayouts')
        bloc_count = 0
        if os.path.exists(layout_dir):
            if progress_callback: progress_callback(2, "Suppression de blocs XML...")
            bloc_count = supprimer_blocs_pic_preencoded(layout_dir)

        if progress_callback: progress_callback(3, "Recompression du PPTX...")
        with zipfile.ZipFile(fichier_sortie, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    full = os.path.join(root, file)
                    rel = os.path.relpath(full, temp_dir)
                    zip_out.write(full, rel)

        if progress_callback: progress_callback(4, "Nettoyage des fichiers temporaires...")
        shutil.rmtree(temp_dir)

        return bloc_count

    except Exception as e:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        raise

st.set_page_config(page_title="Nettoyeur de filigramme Gamma", page_icon="🧹")
st.title("🧹 Nettoyeur de filigramme Gamma (.pptx)")

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

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_out_file:
                fichier_sortie_path = tmp_out_file.name

            def maj_progression(step, message):
                progress_bar.progress(step / 4)
                status.info(message)

            bloc_count = traiter_pptx(
                tmp_file_path,
                fichier_sortie_path,
                maj_progression
            )

            with open(fichier_sortie_path, "rb") as f:
                st.download_button(
                    label=f"📥 Télécharger ({bloc_count} bloc(s) supprimé(s))",
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
