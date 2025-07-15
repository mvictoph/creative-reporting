import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import os
import io
from PIL import Image
import warnings
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

# Ignorer les avertissements OpenPyXL
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Configuration de la page
st.set_page_config(page_title="Amazon Creative Reporting", page_icon="📊", layout="wide")

# Initialisation du client Slack
slack_token = os.environ.get('SLACK_TOKEN')
if slack_token:
    client = WebClient(token=slack_token)
else:
    st.warning("Slack token not found. Slack integration will not work.")
    client = None

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
    return width, height

def calculate_metrics(df, variant):
    variant_data = df[df['Variant'] == variant]
    total_clicks = variant_data['Click-throughs'].sum()
    total_impressions = variant_data['Impressions'].sum()
    ctr = total_clicks / total_impressions if total_impressions > 0 else 0
    return total_clicks, total_impressions, ctr

def apply_amazon_style(text_frame):
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.name = "Amazon Ember"
            run.font.color.rgb = RGBColor(0, 0, 0)

def find_best_matching_image(variant_name, image_files, similarity_threshold=0.6):
    def clean_string(s):
        return ''.join(c.lower() for c in s if c.isalnum())
    
    clean_variant = clean_string(variant_name)
    best_match = None
    highest_similarity = 0
    
    for img_name in image_files:
        clean_img_name = clean_string(os.path.splitext(img_name)[0])
        
        if clean_img_name in clean_variant or clean_variant in clean_img_name:
            similarity = 0.8
        else:
            variant_words = set(clean_variant.split())
            img_words = set(clean_img_name.split())
            common_words = variant_words.intersection(img_words)
            
            if common_words:
                similarity = len(common_words) / max(len(variant_words), len(img_words))
            else:
                common_chars = set(clean_img_name) & set(clean_variant)
                similarity = len(common_chars) / len(set(clean_img_name + clean_variant))
        
        if similarity > highest_similarity:
            highest_similarity = similarity
            best_match = img_name
    
    return (best_match, highest_similarity) if highest_similarity >= similarity_threshold else (None, 0)

def create_ppt_from_data(df, images_dict):
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    
    # Titre - Position ajustée
    title_box = slide.shapes.add_textbox(Inches(1), Inches(0.2), Inches(8), Inches(0.5))
    title_frame = title_box.text_frame
    title_frame.text = "Creative Reporting"
    title_frame.paragraphs[0].font.size = Pt(24)
    apply_amazon_style(title_frame)
    
    # Configuration de la grille
    items_per_row = 2
    left_margin = Inches(1)
    top_margin = Inches(0.7)
    max_image_width = pixels_to_inches(220)
    max_image_height = pixels_to_inches(180)
    spacing_x = max_image_width + Inches(2)
    spacing_y = max_image_height + Inches(1.3)
    
    for index, variant in enumerate(df['Variant'].unique()):
        row_num = index // items_per_row
        col_num = index % items_per_row
        left = left_margin + (col_num * spacing_x)
        top = top_margin + (row_num * spacing_y)
        
        try:
            # Ajout de l'image
            matched_image, similarity = find_best_matching_image(variant, images_dict.keys())
            if matched_image:
                img = images_dict[matched_image]
                width, height = resize_image(img, 220, 180)
                image_width = pixels_to_inches(width)
                image_height = pixels_to_inches(height)
                
                img_byte_arr = io.BytesIO()
                img.save(img_byte_arr, format=img.format if img.format else 'JPEG')
                img_byte_arr.seek(0)
                
                slide.shapes.add_picture(img_byte_arr, left, top, width=image_width, height=image_height)
            
            # Ajout des informations textuelles
            text_top = top + image_height + Inches(0.1)
            
            # Texte Creative simplifié
            name_box = slide.shapes.add_textbox(left, text_top, image_width, Inches(0.2))
            name_frame = name_box.text_frame
            p = name_frame.paragraphs[0]
            p.text = f"Creative: {variant}"
            p.font.size = Pt(9)
            p.font.bold = True
            apply_amazon_style(name_frame)
            
            # Métriques
            total_clicks, total_impressions, ctr = calculate_metrics(df, variant)
            metrics_box = slide.shapes.add_textbox(left, text_top + Inches(0.25), image_width, Inches(0.6))
            metrics_frame = metrics_box.text_frame
            
            metrics = [
                f"Click-throughs: {total_clicks:,}",
                f"Impressions: {total_impressions:,}",
                f"CTR: {ctr:.2%}"
            ]
            
            for idx, metric in enumerate(metrics):
                if idx == 0:
                    p = metrics_frame.paragraphs[0]
                else:
                    p = metrics_frame.add_paragraph()
                p.text = metric
                p.font.size = Pt(9)
            
            apply_amazon_style(metrics_frame)
            
        except Exception as e:
            st.error(f"Error with variant {variant}: {str(e)}")
    
    pptx_buffer = io.BytesIO()
    prs.save(pptx_buffer)
    pptx_buffer.seek(0)
    return pptx_buffer

def send_report_to_slack(file_buffer, filename, user_login):
    if not client:
        return False, "Slack client not initialized. Check your SLACK_TOKEN."
    
    try:
        # Personnaliser le message avec le login de l'utilisateur
        message = f"Here's the latest report from {user_login}!"

        # Upload file directly to channel
        response = client.files_upload_v2(
            channel="C095U79QZDL",
            file=file_buffer,
            filename=filename,
            initial_comment=message
        )
        
        if response['ok']:
            return True, f"Report successfully sent to Slack for {user_login}!"
        else:
            return False, f"Error in Slack response: {response}"
            
    except SlackApiError as e:
        return False, f"Error sending to Slack: {str(e)}"

def main():
    st.title("Creative Reporting Generator")
    st.markdown("---")
    
    # Nouvelle section pour le nom du rapport
    report_name = st.text_input(
        "Creative Reporting Name",
        placeholder="e.g., Creative_Reporting_Q1_2024",
        help="Choose a name for your report (optional)",
        key="report_name"
    )
    
    # Si aucun nom n'est fourni, utiliser le nom par défaut
    final_report_name = report_name if report_name else "Creative_Reporting"
    
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📊 Excel File")
        excel_file = st.file_uploader("Choose your Excel file", type=['xlsx'])
        if excel_file:
            st.success("✅ Excel file successfully loaded")

    with col2:
        st.subheader("🖼️ Creative Images")
        image_files = st.file_uploader("Choose your images", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)
        if image_files:
            st.success(f"✅ {len(image_files)} images loaded")

    if excel_file and image_files:
        try:
            df = pd.read_excel(excel_file)
            images_dict = {img_file.name: Image.open(img_file) for img_file in image_files}
            
            st.write(f"Number of variants detected: {len(df['Variant'].unique())}")
            
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("🚀 Generate PowerPoint Report"):
                    with st.spinner('Generating report...'):
                        pptx_buffer = create_ppt_from_data(df, images_dict)
                        st.success("Report generated successfully!")
                        
                        st.download_button(
                            label="📥 Download PowerPoint Report",
                            data=pptx_buffer,
                            file_name=f"{final_report_name}.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )

            with col2:
                with st.expander("📤 Send to Slack (Optional)"):
                    user_login = st.text_input(
                        "Enter your Amazon login to send via Slack",
                        placeholder="e.g., jsmith",
                        help="Optional: Enter your login to send the report to Slack"
                    )
                    
                    if user_login:
                        if st.button("📤 Generate and Send to Slack"):
                            with st.spinner('Generating and sending to Slack...'):
                                pptx_buffer = create_ppt_from_data(df, images_dict)
                                success, message = send_report_to_slack(
                                    pptx_buffer,
                                    f"{final_report_name}_{user_login}.pptx",
                                    user_login
                                )
                                
                                if success:
                                    st.success(message)
                                else:
                                    st.error(message)
                
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
