import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import os
import io
from PIL import Image
import warnings
import zipfile

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def pixels_to_inches(pixels):
    return Inches(pixels / 96)

def resize_image(img, max_width, max_height):
    width, height = img.size
    aspect_ratio = width / height
    
    if width > max_width or height > max_height:
        if aspect_ratio > 1:
            new_width = max_width
            new_height = int(new_width / aspect_ratio)
        else:
            new_height = max_height
            new_width = int(new_height * aspect_ratio)
        
        return new_width, new_height
    else:
        return width, height

def calculate_average_ctr(df):
    avg_ctrs = df.groupby('Variant')['CTR'].mean()
    return avg_ctrs.to_dict()

def apply_amazon_style(text_frame):
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.name = "Amazon Ember"
            run.font.color.rgb = RGBColor(0, 0, 0)

def create_ppt_from_data(df, images_dict):
    # Créer une nouvelle présentation
    prs = Presentation()
    
    # Calculer les CTR moyens
    avg_ctrs = calculate_average_ctr(df)
    
    # Ajouter une seule diapositive
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    
    # Ajouter un titre
    title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(0.5))
    title_frame = title_box.text_frame
    title_frame.text = "Creative Reporting"
    title_frame.paragraphs[0].font.size = Pt(24)
    apply_amazon_style(title_frame)
    
    # Définir la grille
    items_per_row = 2
    left_margin = Inches(1)
    top_margin = Inches(1.5)
    max_image_width = pixels_to_inches(220)
    max_image_height = pixels_to_inches(180)
    spacing_x = max_image_width + Inches(2)
    spacing_y = max_image_height + Inches(0.8)
    
    # Pour chaque variant
    for index, (variant, ctr) in enumerate(avg_ctrs.items()):
        row_num = index // items_per_row
        col_num = index % items_per_row
        
        left = left_margin + (col_num * spacing_x)
        top = top_margin + (row_num * spacing_y)
        
        try:
            # Ajouter l'image si elle existe
            if variant in images_dict:
                img = images_dict[variant]
                width, height = resize_image(img, 220, 180)
                image_width = pixels_to_inches(width)
                image_height = pixels_to_inches(height)
                
                # Convertir l'image PIL en bytes pour pptx
                img_byte_arr = io.BytesIO()
                img.save(img_byte_arr, format=img.format if img.format else 'JPEG')
                img_byte_arr.seek(0)
                
                slide.shapes.add_picture(img_byte_arr, left, top, width=image_width, height=image_height)
            
            # Ajouter le nom de la créative
            creative_box = slide.shapes.add_textbox(left, top + image_height, image_width, Inches(0.2))
            creative_frame = creative_box.text_frame
            creative_frame.text = variant
            creative_frame.paragraphs[0].font.size = Pt(9)
            apply_amazon_style(creative_frame)
            
            # Ajouter le CTR
            ctr_box = slide.shapes.add_textbox(left, top + image_height + Inches(0.2), image_width, Inches(0.2))
            ctr_frame = ctr_box.text_frame
            ctr_frame.text = f"CTR moyen: {ctr:.2%}"
            ctr_frame.paragraphs[0].font.size = Pt(9)
            apply_amazon_style(ctr_frame)
            
        except Exception as e:
            st.error(f"Erreur avec le variant {variant}: {str(e)}")
    
    # Sauvegarder la présentation dans un buffer
    pptx_buffer = io.BytesIO()
    prs.save(pptx_buffer)
    pptx_buffer.seek(0)
    
    return pptx_buffer

def main():
    st.title("Générateur de Rapport Creative")
    st.write("Uploadez votre fichier Excel et les images correspondantes")
    
    # Upload du fichier Excel
    excel_file = st.file_uploader("Choisissez votre fichier Excel", type=['xlsx'])
    
    # Upload multiple des images
    image_files = st.file_uploader("Choisissez vos images", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)
    
    if excel_file is not None and image_files:
        try:
            # Lire le fichier Excel
            df = pd.read_excel(excel_file)
            
            # Créer un dictionnaire des images
            images_dict = {}
            for img_file in image_files:
                img = Image.open(img_file)
                # Utiliser le nom du fichier sans extension comme clé
                variant_name = os.path.splitext(img_file.name)[0]
                images_dict[variant_name] = img
            
            # Afficher les informations
            unique_variants = df['Variant'].unique()
            st.write(f"Nombre de variants détectés : {len(unique_variants)}")
            st.write("Variants trouvés :")
            for variant in unique_variants:
                st.write(f"- {variant}")
            
            # Bouton pour générer le PPT
            if st.button("Générer le PowerPoint"):
                pptx_buffer = create_ppt_from_data(df, images_dict)
                
                # Proposer le téléchargement
                st.download_button(
                    label="Télécharger le PowerPoint",
                    data=pptx_buffer,
                    file_name="Creative_Reporting.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                
        except Exception as e:
            st.error(f"Une erreur s'est produite : {str(e)}")

if __name__ == "__main__":
    main()
