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

# Définir les colonnes requises pour chaque type de rapport
CTR_REQUIRED_COLUMNS = ['Variant', 'Click-throughs', 'Impressions']
VCR_REQUIRED_COLUMNS = ['Variant', 'Video start', 'Video complete']

def validate_columns(df, report_type):
    required_columns = CTR_REQUIRED_COLUMNS if report_type == "CTR Report" else VCR_REQUIRED_COLUMNS
    missing_columns = [col for col in required_columns if col not in df.columns]
    return len(missing_columns) == 0, missing_columns

def calculate_metrics(df, variant, report_type):
    variant_data = df[df['Variant'] == variant]
    if report_type == "CTR Report":
        total_clicks = variant_data['Click-throughs'].sum()
        total_impressions = variant_data['Impressions'].sum()
        rate = total_clicks / total_impressions if total_impressions > 0 else 0
        return total_clicks, total_impressions, rate
    else:  # VCR Report
        video_starts = variant_data['Video start'].sum()
        video_completes = variant_data['Video complete'].sum()
        vcr = video_completes / video_starts if video_starts > 0 else 0
        return video_starts, video_completes, vcr

def create_ppt_from_data(df, images_dict, report_type):
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    
    # Titre adapté selon le type de rapport
    title_box = slide.shapes.add_textbox(Inches(1), Inches(0.2), Inches(8), Inches(0.5))
    title_frame = title_box.text_frame
    title_frame.text = f"{'Creative Performance Report' if report_type == 'CTR Report' else 'Video Performance Report'}"
    title_frame.paragraphs[0].font.size = Pt(24)
    apply_amazon_style(title_frame)
    
    # Configuration de la grille
    items_per_row = 2
    left_margin = Inches(1)
    top_margin = Inches(1.0)
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
            
            # Texte Creative
            name_box = slide.shapes.add_textbox(left, text_top, image_width, Inches(0.2))
            name_frame = name_box.text_frame
            p = name_frame.paragraphs[0]
            p.text = f"Creative: {variant}"
            p.font.size = Pt(9)
            p.font.bold = True
            apply_amazon_style(name_frame)
            
            # Métriques adaptées selon le type de rapport
            val1, val2, rate = calculate_metrics(df, variant, report_type)
            metrics_box = slide.shapes.add_textbox(left, text_top + Inches(0.25), image_width, Inches(0.6))
            metrics_frame = metrics_box.text_frame
            
            if report_type == "CTR Report":
                metrics = [
                    f"Click-throughs: {val1:,}",
                    f"Impressions: {val2:,}",
                    f"CTR: {rate:.2%}"
                ]
            else:
                metrics = [
                    f"Video starts: {val1:,}",
                    f"Video completes: {val2:,}",
                    f"VCR: {rate:.2%}"
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

def main():
    st.title("Creative Reporting Generator")
    st.markdown("---")
    
    # Sélection du type de rapport
    report_type = st.radio(
        "Select Report Type:",
        ("CTR Report", "VCR Report"),
        help="Choose CTR for static creatives or VCR for video creatives"
    )

    # Nouvelle section pour le nom du rapport
    report_name = st.text_input(
        "Creative Reporting Name",
           placeholder="e.g. Google_Creative_Report_Q1_2024",
        help="Choose a name for your report (optional)",
        key="report_name"
    )
    
    # Si aucun nom n'est fourni, utiliser le nom par défaut
    final_report_name = report_name if report_name else "Creative_Reporting"
    
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📊 Excel File")
        excel_file = st.file_uploader("Upload your Excel file", type=['xlsx'])
        if excel_file:
            st.success("✅ Excel file successfully loaded")

    with col2:
        st.subheader("🖼️ Creative Images")
        image_files = st.file_uploader("Upload your creative images", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)
        if image_files:
            st.success(f"✅ {len(image_files)} images loaded")

    if excel_file and image_files:
        try:
            df = pd.read_excel(excel_file)
            
            # Valider les colonnes
            columns_valid, missing_columns = validate_columns(df, report_type)
            if not columns_valid:
                st.error(f"Missing required columns for {report_type}: {', '.join(missing_columns)}")
                return

            images_dict = {img_file.name: Image.open(img_file) for img_file in image_files}
            
            st.write(f"Number of variants detected: {len(df['Variant'].unique())}")
            
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("🚀 Generate PowerPoint Report"):
                    with st.spinner('Generating report...'):
                        pptx_buffer = create_ppt_from_data(df, images_dict, report_type)
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
                        placeholder="e.g. jsmith",
                        help="Optional: Enter your login to send the report to Slack"
                    )
                    
                    if user_login:
                        if st.button("📤 Generate and Send to Slack"):
                            with st.spinner('Generating and sending to Slack...'):
                                pptx_buffer = create_ppt_from_data(df, images_dict, report_type)
                                success, message = send_report_to_slack(
                                    pptx_buffer,
                                    f"{final_report_name}.pptx",
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
