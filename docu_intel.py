import streamlit as st  
from pptx import Presentation  
from pptx.enum.shapes import MSO_SHAPE_TYPE  
from PIL import Image  
from io import BytesIO  
import requests  
import base64  
from docx import Document  
from docx.shared import Pt  
import fitz  # PyMuPDF  
import os  
import cv2  
import numpy as np  
  
# Azure OpenAI credentials  
azure_endpoint = "https://gpt-4omniwithimages.openai.azure.com/"  
api_key = "6e98566acaf24997baa39039b6e6d183"  
api_version = "2024-02-01"  
model = "GPT-40-mini"  
  
# URL of your Azure function endpoint  
azure_function_url = 'https://doc2pdf.azurewebsites.net/api/HttpTrigger1'  
  
# Function to convert PPT to PDF using Azure Function  
def ppt_to_pdf(ppt_file, pdf_file):  
    mime_type = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'  
    headers = {  
        "Content-Type": "application/octet-stream",  
        "Content-Type-Actual": mime_type  
    }  
    with open(ppt_file, 'rb') as file:  
        response = requests.post(azure_function_url, data=file.read(), headers=headers)  
        if response.status_code == 200:  
            with open(pdf_file, 'wb') as pdf_out:  
                pdf_out.write(response.content)  
            return True  
        else:  
            st.error(f"File conversion failed with status code: {response.status_code}")  
            st.error(f"Response: {response.text}")  
            return False  
  
# Function to encode image as base64  
def encode_image(image):  
    return base64.b64encode(image).decode("utf-8")  
  
def get_image_explanation(base64_image):  
    headers = {  
        "Content-Type": "application/json",  
        "api-key": api_key  
    }  
    data = {  
        "model": model,  
        "messages": [  
            {"role": "system", "content": "You are a helpful assistant that responds in Markdown."},  
            {"role": "user", "content": [  
                {"type": "text", "text": "Explain the content of this image in a single, coherent paragraph. The explanation should be concise and semantically meaningful, summarizing all major points from the image in one continuous paragraph. Avoid using bullet points or separate lists."},  
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64_image}"}}  
            ]}  
        ],  
        "temperature": 0.7  
    }  
  
    response = requests.post(  
        f"{azure_endpoint}/openai/deployments/{model}/chat/completions?api-version={api_version}",  
        headers=headers,  
        json=data  
    )  
  
    if response.status_code == 200:  
        result = response.json()  
        return result["choices"][0]["message"]["content"]  
    else:  
        st.error(f"Error: {response.status_code} - {response.text}")  
        return None  
  
def extract_text_from_ppt(ppt_file):  
    presentation = Presentation(ppt_file)  
    text_content = []  
    for slide_number, slide in enumerate(presentation.slides, start=1):  
        slide_text = []  
        for shape in slide.shapes:  
            if hasattr(shape, "text"):  
                slide_text.append(shape.text)  
        slide_title = slide.shapes.title.text if slide.shapes.title else "Untitled Slide"  
        text_content.append({"slide_number": slide_number, "slide_title": slide_title, "text": " ".join(slide_text)})  
    return text_content  
  
def identify_visual_elements(ppt_file):  
    presentation = Presentation(ppt_file)  
    visual_slides = []  
    for slide_number, slide in enumerate(presentation.slides, start=1):  
        has_visual_elements = False  
        for shape in slide.shapes:  
            if shape.shape_type in {MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.TABLE, MSO_SHAPE_TYPE.CHART, MSO_SHAPE_TYPE.GROUP, MSO_SHAPE_TYPE.AUTO_SHAPE}:  
                has_visual_elements = True  
                break  
        if has_visual_elements:  
            visual_slides.append(slide_number)  
    return visual_slides  
  
def capture_slide_images(pdf_file, slide_numbers):  
    doc = fitz.open(pdf_file)  
    images = []  
    for slide_number in slide_numbers:  
        page = doc[slide_number - 1]  
        pix = page.get_pixmap()  
        image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)  
        buffer = BytesIO()  
        image.save(buffer, format="PNG")  
        images.append({"slide_number": slide_number, "image": buffer.getvalue()})  
    return images  
  
def generate_text_insights(text_content, visual_slides, text_length):  
    headers = {  
        "Content-Type": "application/json",  
        "api-key": api_key  
    }  
    insights = []  
  
    # Set temperature based on text_length  
    if text_length == "Standard":  
        temperature = 0.3  
    elif text_length == "Blend":  
        temperature = 0.5  
    elif text_length == "Creative":  
        temperature = 0.7  
  
    for slide in text_content:  
        slide_text = slide['text']  
        slide_number = slide['slide_number']  
        if len(slide_text.split()) < 20 and slide_number not in visual_slides:  
            continue  # Skip slides with fewer than 20 words and no visual elements  
        prompt = f"""  
     I want you to immediately begin with one of the following phrases based on the slide title:
(a) If the slide title contains the keyword "Background," begin your explanation with "the prior solutions includes..." and proceed by discussing the prior solution presented in the slide. After discussing the prior solution, move forward with the remaining explanation for the slide.
(b) If the slide title contains the keyword "Proposal," start your explanation with "the present disclosure includes..." and proceed by discussing the proposal or invention in the slide. Continue with the remaining explanation.
(c) If the slide title does not contain either of these keywords, begin your explanation with "Aspects of the present disclosure may include..." and discuss the key aspects presented in the slide before moving forward with the rest of the explanation. 
The information should be delivered directly and engagingly in a single, coherent paragraph. Avoid phrases like 'The slide presents,' 'discusses,' 'outlines,' or 'content.' The explanation should be concise and semantically meaningful, summarizing all major points in one paragraph without bullet points. The text should adhere to the following style guidelines:  
        1. Remove all listed profanity words.  
        2. Use passive voice.  
        3. Use conditional and tentative language, such as "may include," "in some aspects," and "aspects of the present disclosure."  
        4. Replace "Million" with "1,000,000" and "Billion" with "1,000,000,000."  
        5. Maintain the following tone characteristics: Precision and Specificity, Formality, Complexity, Objective and Impersonal, Structured and Systematic.  
        6. Follow these style elements: Formal and Objective, Structured and Systematic, Technical Jargon and Terminology, Detailed and Specific, Impersonal Tone, Instructional and Descriptive, Use of Figures and Flowcharts, Legal and Protective Language, Repetitive and Redundant, Examples and Clauses.  
        7. Use the following conditional and tentative language phrases: may include, in some aspects, aspects of the present disclosure, wireless communication networks, by way of example, may be, may further include, may be used, may occur, may use, may monitor, may periodically wake up, may demodulate, may consume, can be performed, may enter and remain, may correspond to, may also include, may be identified in response to, may be further a function of, may be multiplied by, may schedule, may select, may also double, may further comprise, may be configured to, may correspond to a duration value, may correspond to a product of, may be closer, may be significant, may not be able, may result, may reduce, may be operating in, may further be configured to, may further process, may be executed by, may be received, may avoid, may indicate, may be selected, may be proactive, may perform, may be necessary, may be amplified, may involve, may require, may be stored, may be accessed, may be transferred, may be implemented, may include instructions to, may depend upon, may communicate, may be generated, may be configured.  
        8. Maintain the exact wording in the generated content. Do not substitute words with synonyms. For example, "instead" should remain "instead" and not be replaced with "conversely."  
        9. Replace the phrase "further development" with "our disclosure" in all generated content.  
        10. Make sure to use LaTeX formatting for all mathematical symbols, equations, subscripting, and superscripting to ensure they are displayed correctly in the output.  
        11. When encountering programmatic terms or equations, ensure they are accurately represented and contextually retained.  
        {slide_text}  
        """  
        if text_length == "Standard":  
            prompt += "\n\nGenerate a short paragraph."  
        elif text_length == "Blend":  
            prompt += "\n\nGenerate a medium-length paragraph."  
        elif text_length == "Creative":  
            prompt += "\n\nGenerate a longer paragraph."  
  
        data = {  
            "model": model,  
            "messages": [{"role": "system", "content": "You are a helpful assistant."}, {"role": "user", "content": prompt}],  
            "temperature": temperature  
        }  
  
        response = requests.post(  
            f"{azure_endpoint}/openai/deployments/{model}/chat/completions?api-version={api_version}",  
            headers=headers,  
            json=data  
        )  
  
        if response.status_code == 200:  
            result = response.json()  
            insights.append({"slide_number": slide['slide_number'], "slide_title": slide['slide_title'], "insight": result["choices"][0]["message"]["content"]})  
        else:  
            st.error(f"Error: {response.status_code} - {response.text}")  
            insights.append({"slide_number": slide['slide_number'], "slide_title": slide['slide_title'], "insight": "Error in generating insight"})  
  
    return insights  
  
def generate_image_insights(image_content, text_length, api_key, azure_endpoint, model, api_version):  
    insights = []  
  
    # Set temperature based on text_length  
    if text_length == "Standard":  
        temperature = 0.3  
    elif text_length == "Blend":  
        temperature = 0.5  
    elif text_length == "Creative":  
        temperature = 0.7  
  
    for image_data in image_content:  
        base64_image = encode_image(image_data['image'])  
        slide_number = image_data['slide_number']  
          
        # Determine how many images are on the slide  
        images_on_slide = [img for img in image_content if img['slide_number'] == slide_number]  
        image_ref = f"figure {slide_number}"  
        if len(images_on_slide) > 1:  
            image_ref += f"({chr(97 + images_on_slide.index(image_data))})"  
  
        headers = {  
            "Content-Type": "application/json",  
            "api-key": api_key  
        }  
  
        prompt = f"""  

Start by listing all images present in the slide using the format: 'Referring to {image_ref}(a), {image_ref}(b), etc.' based on the number of images in the slide. If there is only one image, simply use 'Referring to {image_ref},' If there are multiple images, ensure the last reference is separated by 'and,' e.g., '{image_ref}(a) and {image_ref}(b),' for two images. Continue this pattern for slides with more images. After listing the figures, follow the steps below:

Step-1:
After listing the slide reference, begin immediately after the comma with one of the following phrases based on the slide title. Ensure that the word directly following the comma starts with a lowercase letter. This rule must be followed consistently for all slides.
(a) If the slide title contains the keyword "Background," begin your explanation with "the prior solutions includes..." and proceed by discussing the prior solution presented in the slide. After discussing the prior solution, move forward with the remaining explanation for the slide.
(b) If the slide title contains the keyword "Proposal," start your explanation with "the present disclosure includes..." and proceed by discussing the proposal or invention in the slide. Continue with the remaining explanation.
(c) If the slide title does not contain either of these keywords, begin your explanation with "Aspects of the present disclosure include..." and discuss the key aspects presented in the slide before moving forward with the rest of the explanation.

Step-2:
Strictly avoid beginning or using phrases like "The slide" during the explanation to maintain a more natural flow.

Step-3:
Strictly avoid beginning or using phrases like "The text" during the explanation to maintain a more natural flow.

Step-4:
Strictly avoid beginning or using phrases like "The image" during the explanation to maintain a more natural flow.

Step-5:
After referencing the figure, always start the following sentence with "In this aspect," and continue with the detailed explanation of the content.

Step-6:
Start by analyzing the text content of the slide. Reproduce the text as accurately as possible, maintaining the context, and then describe the image. Ensure the explanation smoothly integrates both the text and image content.

Step-7:
Ensure that all referenced images are clearly explained. After referencing the image (e.g., "Referring to {image_ref}(a)," etc.), provide a detailed explanation for each image in the slide.

Step-8:
While explaining, ensure that you follow the style guide:
(a) Remove all listed profanity words.
(b) Use passive voice consistently throughout the explanation.
(c) Avoid using the term "consist" or any form of that verb when describing inventions or disclosures.
(d) Replace "Million" with "1,000,000" and "Billion" with "1,000,000,000."
(e) Maintain precision, specificity, and formality in tone. The explanation should be complex, objective, and structured systematically.
(f) Use technical jargon and terminology that is detailed and specific. Maintain an impersonal tone.
(g) Structure the explanation systematically, and use terms like "defined as," "the first set," "the second set," and "for example."
(h) Use conditional and tentative language such as "may include," "in some aspects," "aspects of the present disclosure," "wireless communication networks," "by way of example," "may be," "may further include," "may be used," "may occur," and other similar phrases.
(i) Capture all key wording and phrases accurately. Do not substitute words with synonyms (e.g., maintain "instead" rather than replacing it with "conversely").
(j) Avoid repeating abbreviations if they have already been defined earlier in the explanation.
(k) When discussing the current disclosure, use definitive language.
(l) Ensure accurate representation and contextual integration of any figures, flowcharts, or equations referenced in the slide.
(m) If the slide or image contains the word "example" or "E.g." ensure it is strictly reproduced exactly as it appears. Apply the same accuracy for all examples present in the image.

Step-9:
I expect you to provide a clear and consistent explanation based on the image. There's no need to mention the steps you're following, use unnecessary formatting (such as bold text), or include unrelated details, topics, or subtopics in your response. Focus solely on delivering a straightforward, cohesive explanation, without describing your process or referring to the current step. Just provide the explanation—nothing more.
"""  
  
        data = {  
            "model": model,  
            "messages": [  
                {"role": "system", "content": "You are a helpful assistant that responds in Markdown."},  
                {"role": "user", "content": [  
                    {"type": "text", "text": prompt},  
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64_image}"}}  
                ]}  
            ],  
            "temperature": temperature  
        }  
  
        response = requests.post(  
            f"{azure_endpoint}/openai/deployments/{model}/chat/completions?api-version={api_version}",  
            headers=headers,  
            json=data  
        )  
  
        if response.status_code == 200:  
            result = response.json()  
            insights.append({  
                "slide_number": image_data['slide_number'],  
                "slide_title": image_data.get('slide_title', 'Untitled Slide'),  
                "insight": result["choices"][0]["message"]["content"]  
            })  
        else:  
            st.error(f"Error: {response.status_code} - {response.text}")  
            insights.append({  
                "slide_number": image_data['slide_number'],  
                "slide_title": image_data.get('slide_title', 'Untitled Slide'),  
                "insight": "Error in generating insight"  
            })  
  
    return insights  
  
def aggregate_content(text_insights, image_insights):  
    aggregated_content = []  
    for text in text_insights:  
        slide_number = text['slide_number']  
        slide_title = text['slide_title']  
        text_insight = text['insight']  
        image_insight = next((img['insight'] for img in image_insights if img['slide_number'] == slide_number), None)  
        if image_insight:  
            aggregated_content.append({  
                "slide_number": slide_number,  
                "slide_title": slide_title,  
                "content": f"{image_insight}"  
            })  
        else:  
            aggregated_content.append({  
                "slide_number": slide_number,  
                "slide_title": slide_title,  
                "content": text_insight  
            })  
    return aggregated_content  
  
def sanitize_text(text):  
    if text:  
        sanitized = ''.join(c for c in text if c.isprintable() and c not in {'\x00', '\x01', '\x02', '\x03', '\x04', '\x05', '\x06', '\x07', '\x08', '\x0B', '\x0C', '\x0E', '\x0F', '\x10', '\x11', '\x12', '\x13', '\x14', '\x15', '\x16', '\x17', '\x18', '\x19', '\x1A', '\x1B', '\x1C', '\x1D', '\x1E', '\x1F'})  
        return sanitized  
    return text  
  
def save_content_to_word(aggregated_content, output_file_name, extracted_images):  
    doc = Document()  
    style = doc.styles['Normal']  
    font = style.font  
    font.name = 'Times New Roman'  
    font.size = Pt(10.5)  # Reduced font size for paragraphs  
    paragraph_format = style.paragraph_format  
    paragraph_format.line_spacing = 1.5  
    paragraph_format.alignment = 3  # Justify  
  
    for slide in aggregated_content:  
        sanitized_title = sanitize_text(slide['slide_title'])  
        sanitized_content = sanitize_text(slide['content'])  
        doc.add_heading(f"[[{slide['slide_number']}, {sanitized_title}]]", level=1)  
        if sanitized_content:  # Only add content if it exists  
            doc.add_paragraph(sanitized_content)  
  
    # Add extracted images after the generated content  
    if extracted_images:  
        doc.add_heading("Extracted Images", level=1)  
        for idx, (image, slide_number) in enumerate(extracted_images):  
            _, buffer = cv2.imencode('.png', image)  
            image_stream = BytesIO(buffer)  
            doc.add_paragraph(f"Image from Slide {slide_number}:")  
            doc.add_picture(image_stream, width=doc.sections[0].page_width - doc.sections[0].left_margin - doc.sections[0].right_margin)  
            doc.add_paragraph("\n")  # Add space after image  
  
    output = BytesIO()  
    doc.save(output)  
    output.seek(0)  
    return output  
  
def extract_and_clean_page_image(page, top_mask, bottom_mask, left_mask, right_mask):  
    # Get the page as an image  
    pix = page.get_pixmap()  
    img_data = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)  
  
    # Convert the image to BGR format for OpenCV  
    img_bgr = cv2.cvtColor(img_data, cv2.COLOR_RGB2BGR)  
  
    # Convert to grayscale for processing  
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)  
  
    # Threshold the image to get binary image  
    _, binary = cv2.threshold(gray, 240, 255, cv2.THRESH_BINARY_INV)  
  
    # Detect contours of possible images/diagrams  
    contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)  
  
    # Check if there are any valid contours (image regions)  
    valid_contours = [cv2.boundingRect(contour) for contour in contours if cv2.boundingRect(contour)[2] > 50 and cv2.boundingRect(contour)[3] > 50]  
    if not valid_contours:  
        return None  # Skip the page if no valid images/diagrams are found  
  
    # Create a mask for the detected contours  
    mask = np.zeros_like(binary)  
    for x, y, w, h in valid_contours:  
        # Apply the adjustable top, bottom, left, and right masking values from the sliders  
        # Ensure coordinates do not go out of image bounds  
        x1 = max(0, x - left_mask)  
        y1 = max(0, y - top_mask)  
        x2 = min(img_bgr.shape[1], x + w + right_mask)  
        y2 = min(img_bgr.shape[0], y + h + bottom_mask)  
        cv2.rectangle(mask, (x1, y1), (x2, y2), 255, -1)  
  
    # Use the mask to keep only the regions with images/diagrams  
    text_removed = cv2.bitwise_and(img_bgr, img_bgr, mask=mask)  
  
    # Set the background to white where the mask is not applied  
    background = np.ones_like(img_bgr) * 255  
    cleaned_image = np.where(mask[:, :, None] == 255, text_removed, background)  
  
    # Convert cleaned image to grayscale  
    cleaned_image_gray = cv2.cvtColor(cleaned_image, cv2.COLOR_BGR2GRAY)  
    return cleaned_image_gray  
  
def extract_images_from_pdf(pdf_file, top_mask, bottom_mask, left_mask, right_mask):  
    # Open the PDF file  
    pdf_document = fitz.open(pdf_file)  
    page_images = []  
  
    for page_num in range(len(pdf_document)):  
        page = pdf_document.load_page(page_num)  
  
        # Extract and clean the page image  
        cleaned_image = extract_and_clean_page_image(page, top_mask, bottom_mask, left_mask, right_mask)  
        if cleaned_image is not None:  
            page_images.append((cleaned_image, page_num + 1))  # Keep track of the slide number  
  
    pdf_document.close()  
    return page_images  
  
def main():  
    st.title("PPT Insights Extractor")  
  
    text_length = st.select_slider(  
        "Content Generation Slider",  
        options=["Standard", "Blend", "Creative"],  
        value="Blend"  
    )  
  
    # Add Title and Information Button for Image Extraction Slider  
    st.sidebar.markdown("### Image Extraction Slider")  
  
    # Initialize session state variables for the sliders  
    if 'top_mask' not in st.session_state:  
        st.session_state.top_mask = 40  
    if 'bottom_mask' not in st.session_state:  
        st.session_state.bottom_mask = 40  
    if 'left_mask' not in st.session_state:  
        st.session_state.left_mask = 85  
    if 'right_mask' not in st.session_state:  
        st.session_state.right_mask = 85  
  
    # Arrange the buttons in a row using columns  
    col1, col2 = st.sidebar.columns(2)  
    with col1:  
        if st.button("Default"):  
            st.session_state.top_mask = 40  
            st.session_state.bottom_mask = 40  
            st.session_state.left_mask = 85  
            st.session_state.right_mask = 85  
  
    with col2:  
        if st.button("A4"):  
            st.session_state.top_mask = 70  
            st.session_state.bottom_mask = 70  
            st.session_state.left_mask = 85  
            st.session_state.right_mask = 85  
  
    # Add sliders to adjust the top, bottom, left, and right masking values  
    top_mask = st.sidebar.slider("Adjust Top Masking Value", min_value=10, max_value=100, value=st.session_state.top_mask, step=1)  
    bottom_mask = st.sidebar.slider("Adjust Bottom Masking Value", min_value=10, max_value=100, value=st.session_state.bottom_mask, step=1)  
    left_mask = st.sidebar.slider("Adjust Left Masking Value", min_value=10, max_value=500, value=st.session_state.left_mask, step=1)  
    right_mask = st.sidebar.slider("Adjust Right Masking Value", min_value=10, max_value=200, value=st.session_state.right_mask, step=1)  
  
    # Update session state if sliders are moved  
    if top_mask != st.session_state.top_mask or bottom_mask != st.session_state.bottom_mask or left_mask != st.session_state.left_mask or right_mask != st.session_state.right_mask:  
        st.session_state.top_mask = top_mask  
        st.session_state.bottom_mask = bottom_mask  
        st.session_state.left_mask = left_mask  
        st.session_state.right_mask = right_mask  
  
    uploaded_ppt = st.file_uploader("Upload a PPT file", type=["pptx"])  
  
    if uploaded_ppt is not None:  
        st.info("Processing PPT file...")  
  
        # Extract the base name of the uploaded PPT file  
        ppt_filename = uploaded_ppt.name  
        base_filename = os.path.splitext(ppt_filename)[0]  
        output_word_filename = f"{base_filename}.docx"  
  
        try:  
            # Convert PPT to PDF  
            with open("temp_ppt.pptx", "wb") as f:  
                f.write(uploaded_ppt.read())  
            if not ppt_to_pdf("temp_ppt.pptx", "temp_pdf.pdf"):  
                st.error("PDF conversion failed. Please check the uploaded PPT file.")  
                return  
  
            # Extract text and identify slides with visual elements  
            text_content = extract_text_from_ppt("temp_ppt.pptx")  
            visual_slides = identify_visual_elements("temp_ppt.pptx")  
  
            # Capture images of marked slides  
            slide_images = capture_slide_images("temp_pdf.pdf", visual_slides)  
  
            st.info("Generating text insights...")  
            text_insights = generate_text_insights(text_content, visual_slides, text_length)  
  
            st.info("Generating image insights...")  
            image_insights = generate_image_insights(slide_images, text_length, api_key, azure_endpoint, model, api_version)  
  
            st.info("Extracting additional images...")  
            extracted_images = extract_images_from_pdf("temp_pdf.pdf", top_mask, bottom_mask, left_mask, right_mask)  
  
            st.info("Aggregating content...")  
            aggregated_content = aggregate_content(text_insights, image_insights)  
  
            st.info("Saving to Word document...")  
            output_doc = save_content_to_word(aggregated_content, output_word_filename, extracted_images)  
  
            st.download_button(label="Download Word Document", data=output_doc, file_name=output_word_filename)  
  
            st.success("Processing completed successfully!")  
        except Exception as e:  
            st.error(f"An error occurred: {e}")  
  
if __name__ == "__main__":  
    main()  

