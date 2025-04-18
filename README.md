
# Nettoyeur de Filigrane Gamma pour fichiers PowerPoint (.pptx)

Cet outil supprime automatiquement certains filigranes (notamment les images avec la description `preencoded.png`) intégrés dans les layouts de fichiers PowerPoint `.pptx`.  
Il dispose d'une interface graphique simple basée sur Tkinter.

## ✨ Fonctionnalités

- Interface graphique intuitive
- Suppression ciblée de blocs `<p:pic>` dans les layouts
- Traitement automatique : décompression → nettoyage → recompression
- Barre de progression et messages de statut

## 🖥️ Prérequis

- Python 3.7+
- Système d'exploitation : Windows, Linux ou macOS

## 📦 Dépendances

Aucune dépendance externe requise (utilise uniquement la bibliothèque standard Python).

## 🚀 Utilisation

1. **Télécharge ou clone ce dépôt :**

   ```bash
   git clone https://github.com/MicheDePainn/Gamma-Watermark-Cleaner-.pptx
   cd nettoyeur-pptx
   ```

2. **Lance l'application :**

   ```bash
   python nettoyeur.py
   ```

3. **Dans l'application :**
   - Sélectionne un fichier `.pptx` d'entrée.
   - Choisis un fichier `.pptx` de sortie.
   - Clique sur **"Lancer le nettoyage"**.
   - À la fin, le nombre de blocs supprimés est affiché.

## 📁 Exemple de fonctionnement

Le programme supprime les blocs XML `<p:pic>` des fichiers dans `ppt/slideLayouts/` lorsque l’attribut `descr` de `<p:cNvPr>` est égal à `preencoded.png`.

```xml
<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="4" name="Image 4" descr="preencoded.png"/>
    ...
  </p:nvPicPr>
  ...
</p:pic>
```

## 📃 Licence

Ce projet est libre et ouvert. Tu peux le modifier et l’utiliser comme bon te semble.

---

**Développé avec ❤️ en Python.**
