import streamlit as st
import os
from pptx import Presentation
from pptx.util import Inches
from PIL import Image
import io

def create_pptx_from_images(images):
    # 新しいプレゼンテーションを作成
    prs = Presentation()
    
    # スライドサイズを16:9に設定（2880×1620ピクセルに対応）
    prs.slide_width = Inches(13.3333)
    prs.slide_height = Inches(7.5)
    
    # 空白のスライドレイアウトを使用
    blank_slide_layout = prs.slide_layouts[6]
    
    # 各画像を新しいスライドに追加
    for uploaded_file in images:
        slide = prs.slides.add_slide(blank_slide_layout)
        # 画像をバイトストリームとして読み込む
        image_stream = io.BytesIO(uploaded_file.getvalue())
        slide.shapes.add_picture(image_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
    
    # プレゼンテーションをバイトストリームとして保存
    pptx_stream = io.BytesIO()
    prs.save(pptx_stream)
    return pptx_stream.getvalue()

def main():
    st.title("Images to PowerPoint Converter")
    st.write("Upload your images and convert them into a PowerPoint presentation.")
    
    # 複数の画像をアップロード
    uploaded_files = st.file_uploader("Choose images", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
    
    if uploaded_files:
        st.write(f"Number of images uploaded: {len(uploaded_files)}")
        
        # プレビューを表示（最初の5枚まで）
        cols = st.columns(min(5, len(uploaded_files)))
        for i, (col, image) in enumerate(zip(cols, uploaded_files[:5])):
            with col:
                st.image(image, caption=f"Image {i+1}", use_column_width=True)
        
        if st.button("Convert to PowerPoint"):
            with st.spinner("Creating PowerPoint presentation..."):
                try:
                    pptx_bytes = create_pptx_from_images(uploaded_files)
                    
                    # ダウンロードボタンを表示
                    st.download_button(
                        label="Download PowerPoint",
                        data=pptx_bytes,
                        file_name="presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                    st.success("PowerPoint presentation created successfully!")
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main() 