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
import io  
from azure.storage.blob import BlobServiceClient, BlobClient, ContainerClient  
from azure.core.exceptions import ResourceExistsError  
import tempfile  
import re
  
# Azure OpenAI credentials  
azure_endpoint = "https://gpt-4omniwithimages.openai.azure.com/"  
api_key = "6e98566acaf24997baa39039b6e6d183"  
api_version = "2024-02-01"  
model = "GPT-40-mini"  
  
# Azure Blob Storage credentials  
connection_string = "DefaultEndpointsProtocol=https;AccountName=patentpptapp;AccountKey=4988gBY4D2RU4zdy1NCUoORdCRYvoOziWSHK9rOVHxy9pFXfKenRqyE/P+tpFpfmNObUm/zOCjeY+AStiCS3uw==;EndpointSuffix=core.windows.net"  
container_name = "ppt-storage"  
  
blob_service_client = BlobServiceClient.from_connection_string(connection_string)  
  
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
                {"type": "text", "text": "Explain the content of this image in a single, coherent paragraph. The explanation should be concise and semantically meaningful, summarizing all major points from the image in one paragraph. Avoid using bullet points or separate lists."},  
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
            if shape.has_text_frame:  
                for paragraph in shape.text_frame.paragraphs:  
                    for run in paragraph.runs:  
                        slide_text.append(run.text)  
        slide_title = slide.shapes.title.text if slide.shapes.title else "Untitled Slide"  
        text_content.append({"slide_number": slide_number, "slide_title": slide_title, "text": " ".join(slide_text)})  
    return text_content  
  
def is_image_of_interest(shape):  
    """Check if a shape contains an image in formats of interest"""  
    try:  
        if hasattr(shape, "image"):  
            image_ext = os.path.splitext(shape.image.filename)[1].lower()  
            if image_ext in [".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff"]:  
                return image_ext  
    except Exception:  
        pass  
    return None  
  
def detect_image_slides(ppt_bytes):  
    """Detect slides containing images in the desired formats"""  
    ppt = Presentation(io.BytesIO(ppt_bytes))  
    image_slides = {}  
    for i, slide in enumerate(ppt.slides):  
        for shape in slide.shapes:  
            image_format = is_image_of_interest(shape)  
            if image_format:  
                slide_number = i + 1  
                image_slides[slide_number] = image_format  
                break  
    return image_slides  
  
def identify_visual_elements(ppt_bytes):  
    """Identify slides with visual elements"""  
    presentation = Presentation(io.BytesIO(ppt_bytes))  
    visual_slides = []  
    for slide_number, slide in enumerate(presentation.slides, start=1):  
        has_visual_elements = False  
        for shape in slide.shapes:  
            if shape.shape_type in {MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.TABLE, MSO_SHAPE_TYPE.CHART,   
                                    MSO_SHAPE_TYPE.GROUP, MSO_SHAPE_TYPE.AUTO_SHAPE}:  
                has_visual_elements = True  
                break  
        if has_visual_elements:  
            visual_slides.append(slide_number)  
    return visual_slides  
  
def combine_slide_numbers(image_slides, visual_slides):  
    """Combine slide numbers from image slides and visual element slides"""  
    combined_slides = set(image_slides.keys()).union(set(visual_slides))  
    return sorted(list(combined_slides))  
  
def capture_slide_images(pdf_file, slide_numbers, low_quality_slides):  
    """Capture images from identified slides in the PDF"""  
    doc = fitz.open(pdf_file)  
    images = []  
    for slide_number in slide_numbers:  
        if slide_number in low_quality_slides:  
            continue  # Skip low-quality slides  
  
        page = doc[slide_number - 1]  
        pix = page.get_pixmap()  
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)  
        buffer = BytesIO()  
        img.save(buffer, format="PNG")  
        images.append({"slide_number": slide_number, "image": buffer.getvalue()})  
    return images  
  
def generate_theme(text_content, image_content):  
    headers = {  
        "Content-Type": "application/json",  
        "api-key": api_key  
    }  
    combined_text = " ".join([slide['text'] for slide in text_content])  
    combined_images = " ".join([get_image_explanation(encode_image(image['image'])) for image in image_content])  
    prompt = f"Generate a high-level cohesive theme that encapsulates the entire content of the PPT (both text and images). The theme should capture key ideas, keywords, and terminology from both the extracted text and images.\n\nText Content: {combined_text}\n\nImage Content: {combined_images}"  
  
    data = {  
        "model": model,  
        "messages": [{"role": "system", "content": "You are a helpful assistant."}, {"role": "user", "content": prompt}],  
        "temperature": 0.5  
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
  
def generate_text_insights(text_content, visual_slides, text_length, theme, low_quality_slides):  
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
        slide_number = slide['slide_number']  
        if slide_number in low_quality_slides:  
            continue  # Skip low-quality slides  
  
        slide_text = slide['text']  
        
        prompt = f"""  
        {theme}  
        I want you to begin with one of the following phrases based on the slide title:  
        
        (a) If and only if the slide title contains the keyword "Background," begin the explanation with "The prior solutions include..." Proceed by discussing only the prior solution presented in the slide. Ensure no mention of any proposal or disclosure occurs at this stage, and strictly limit the explanation to the prior solutions.  
        
        (b) If and only if the slide title contains the keyword "Proposal," start the explanation with "The present disclosure includes..." Focus exclusively on discussing the proposal or invention presented in the slide. Ensure that no background information is referenced, and strictly adhere to the proposal/invention-related content.  
        
        (c) If the slide title does not contain either "Background" or "Proposal," start the explanation with "Aspects of the present disclosure include..." Discuss the key aspects of the slide's content, ensuring no mention of prior solutions or proposals. Adhere to the neutral tone, focusing on the core aspects of the slide's content.  
        
        The information should be delivered in a structured, clear, and concise paragraph while avoiding phrases like 'The slide presents,' 'discusses,' 'outlines,' or 'content.' Summarize all major points without bullet points.  
        
        Follow these detailed style guidelines for the generated content:  
        
        1. Remove all listed profanity words.  
        2. Use passive voice consistently.  
        3. Use conditional and tentative language, such as "may include," "in some aspects," and "aspects of the present disclosure."  
        4. Replace "Million" with "1,000,000" and "Billion" with "1,000,000,000."  
        5. Maintain these tone characteristics: Precision and Specificity, Formality, Complexity, Objectivity and Impersonality, Structured and Systematic.  
        6. Follow these style elements: Formal and Objective, Structured and Systematic, Technical Jargon and Terminology, Detailed and Specific, Impersonal Tone, Instructional and Descriptive, Use of Figures and Flowcharts, Legal and Protective Language, Repetitive and Redundant, Examples and Clauses.  
        7. Use the following conditional and tentative language phrases: may include, in some aspects, aspects of the present disclosure, by way of example, may be, may further include, may be used, may occur, may use, may monitor, may periodically wake up, may demodulate, may consume, can be performed, may enter and remain, may correspond to, may also include, may be identified in response to, may be further a function of, may be multiplied by, may schedule, may select, may also double, may further comprise, may be configured to, may correspond to a duration value, may correspond to a product of, may be closer, may be significant, may not be able, may result, may reduce, may be operating in, may further be configured to, may further process, may be executed by, may be received, may avoid, may indicate, may be selected, may be proactive, may perform, may be necessary, may be amplified, may involve, may require, may be stored, may be accessed, may be transferred, may be implemented, may include instructions to, may depend upon, may communicate, may be generated, may be configured.  
        8. Maintain the exact wording in the generated content. Do not substitute words with synonyms. For example, "instead" should remain "instead" and not be replaced with "conversely."  
        9. Replace the phrase "further development" with "our disclosure" in all generated content.  
        10. Use LaTeX formatting for all mathematical symbols, equations, subscripting, and superscripting to ensure they are displayed correctly in the output.  
        11. Accurately represent and contextually retain programmatic terms or equations.
        12. Avoid expanding abbreviations under any circumstances. Use abbreviations exactly as they appear in the extracted content. If an abbreviation is present, reproduce it as is, without repeating or expanding it at any point throughout the entire explanation.
        13. Strictly avoid using the following words and phrases in your explanations: 'consist,' 'consisting,' 'necessary,' 'explore,' 'exploration,' 'key component,' 'revolutionizing,' 'innovative,' or any similar adjectives.
        {slide_text}  
        """  
  
        if text_length == "Standard":  
            prompt += "\n\nGenerate a short paragraph"  
        elif text_length == "Blend":  
            prompt += "\n\nGenerate a medium-length paragraph"  
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
  
def generate_image_insights(image_content, text_length, api_key, azure_endpoint, model, api_version, theme, low_quality_slides):  
    insights = []  
  
    # Set temperature based on text_length  
    if text_length == "Standard":  
        temperature = 0.3  
    elif text_length == "Blend":  
        temperature = 0.5  
    elif text_length == "Creative":  
        temperature = 0.7  
  
    for image_data in image_content:  
        slide_number = image_data['slide_number']  
        if slide_number in low_quality_slides:  
            continue  # Skip low-quality slides  
  
        base64_image = encode_image(image_data['image'])  
  
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
        {theme}  
        Step-1: Begin by detecting and listing all figures present in the slide, ensuring no figure is overlooked. This includes cases where multiple figures are arranged in parallel, adjacent to each other, or stacked in a sequence. Treat each unique diagram, sketch, or flowchart as a separate figure.

        Step-1(a): Reference all figures explicitly and sequentially using the format:
        “Referring to Figure {image_ref}(a), Figure {image_ref}(b), and Figure {image_ref}(c)” for three figures, or “Referring to Figure {image_ref}(a)” and “Referring to Figure {image_ref}(b)” for two figures.

        Important: For slides containing multiple figures, ensure that each figure is individually referenced and mentioned, regardless of any similarities between them. Each figure reference must be unique and should not be skipped or combined with other figures.
        Step-1(b): Verify that the references are included in the generated output. Each figure must be mentioned before its detailed description. The order of references should match the visual order of the figures as they appear in the slide.

        Step-1(c): After referencing the figures, describe each one individually and in sequence, using the specific reference format consistently. For example:

        “Figure {image_ref}(a) illustrates...”
        “Figure {image_ref}(b) shows...”
        “Figure {image_ref}(c) demonstrates...”
        Ensure each figure’s role and relevance are fully explained without omitting any figure or detail.

        Step-1(d): If the generated output does not reference all detected figures or merges references (e.g., saying “the figures” instead of listing them individually), flag the response and provide a detailed alert noting which figures were missed or incorrectly referenced. This ensures every figure is distinctly acknowledged and described.

        Step-1(e): For slides containing complex or overlapping figures, describe their relationships and interactions clearly. Avoid using generalized terms like “the figures” and instead use precise references (e.g., “Figure {image_ref}(a) interacts with Figure {image_ref}(b)...”).

        Final Check: Ensure the response includes all figures as detected, with each figure uniquely referenced and accurately described. Cross-check the references to ensure they are complete and in the correct order.
        
        Finally, follow the steps below:  
        
        Step-2: After listing the slide reference, begin immediately after the comma with one of the following phrases based on the slide title. Ensure that the word directly following the comma starts with a lowercase letter. This rule must be followed consistently for all slides.  
        
        Step-3: If and only if the slide title contains the keyword "Background," begin the explanation with "The prior solutions include..." Proceed by discussing only the prior solution presented in the slide. Ensure no mention of any proposal or disclosure occurs at this stage, and strictly limit the explanation to the prior solutions.  
        
        Step-4: If and only if the slide title contains the keyword "Proposal," start the explanation with "The present disclosure includes..." Focus exclusively on discussing the proposal or invention presented in the slide. Ensure that no background information is referenced, and strictly adhere to the proposal/invention-related content.  
        
        Step-5: If the slide title does not contain either "Background" or "Proposal," start the explanation with "Aspects of the present disclosure include..." Discuss the key aspects of the slide's content, ensuring no mention of prior solutions or proposals. Adhere to the neutral tone, focusing on the core aspects of the slide's content. Also, every image or slide that contains the word "example" or "e.g.," ensure that the word is reproduced exactly as it appears, every time it is used. Each example must be thoroughly explained, and the word "example" should consistently be used when referring to examples. Avoid using alternative words such as "additionally" or "furthermore." If the image contains multiple examples, ensure that all examples are explained in detail and that each occurrence of the word "example" is properly included in the explanation. No examples should be overlooked or combined into a single sentence without proper reference to each one.
        
        Step-6: For images identified as graphs, provide an explanation that captures the overall meaning of the graph, including a detailed description of the x and y axes.  
        
        Step-7: For every image containing a perspective view, ensure that the perspective is identified and described in detail. Begin by clearly stating that the image has a perspective view, followed by a thorough explanation of the perspective itself, including angles, depth, and spatial relationships within the image. This must be done for every image that features a perspective view without exception. Ensure that no image with a perspective view is overlooked, and the explanation captures the full depth and context of the perspective in a brief but comprehensive manner.  
        
        Step-8: Instead of labeling the images as "left figure" or "right figure" refer to them using a specific reference that identifies which figure is being referenced.  
        
        Step-9: For every image or slide that contains the word "example" or "e.g.," ensure that the word is reproduced exactly as it appears, every time it is used. Each example must be thoroughly explained, and the word "example" should consistently be used when referring to examples. Avoid using alternative words such as "additionally" or "furthermore." If the image contains multiple examples, ensure that all examples are explained in detail and that each occurrence of the word "example" is properly included in the explanation. No examples should be overlooked or combined into a single sentence without proper reference to each one.  
        
        Additionally, I have a slide with the following bullet points:  
        Point 1: [Your first point]  
        Point 2: [Your second point]  
        Point 3: [Your third point, which includes examples Eg 1 and Eg 2]  
        Please generate a cohesive paragraph that integrates these points while ensuring that all examples are specifically called out and fully explained. Always use the word "example" each time it appears in the content, and maintain clarity and continuity in the overall description.  
        
        Step-10: Strictly avoid beginning or using phrases like "The slide" during the explanation to maintain a more natural flow.  
        
        Step-11: Strictly avoid beginning or using phrases like "The text" during the explanation to maintain a more natural flow.  
        
        Step-12: Strictly avoid beginning or using phrases like "The image" during the explanation to maintain a more natural flow.  
        
        Step-13: After referencing the figure, always start the following sentence with "In this aspect," and continue with the detailed explanation of the content.  
        
        Step-14: Start by analyzing the text content of the slide. Reproduce the text as accurately as possible, maintaining the context, and then describe the image. Ensure the explanation smoothly integrates both the text and image content.  
        
        Step-15: While explaining, ensure that you follow the style guide step-by-step from (a) to (j):  
        (a) Remove all listed profanity words.  
        (b) Use passive voice consistently throughout the explanation.  
        (c) Replace "Million" with "1,000,000" and "Billion" with "1,000,000,000."  
        (d) Maintain precision, specificity, and formality in tone. The explanation should be complex, objective, and structured systematically.  
        (e) Use technical jargon and terminology that is detailed and specific. Maintain an impersonal tone.  
        (f) Structure the explanation systematically, and use terms like "defined as," "the first set," "the second set," and "for example."  
        (g) Use conditional and tentative language such as "may include," "in some aspects," "aspects of the present disclosure," "by way of example," "may be," "may further include," "may be used," "may occur," and other similar phrases.  
        (h) Capture all key wording and phrases accurately. Do not substitute words with synonyms (e.g., maintain "instead" rather than replacing it with "conversely").  
        (i) When discussing the current disclosure, use definitive language.  
        (j) Ensure accurate representation and contextual integration of any figures, flowcharts, or equations referenced in the slide.  
        (k) Strictly avoid using the following words and phrases in your explanations: 'consist,' 'consisting,' 'necessary,' 'explore,' 'exploration,' 'key component,' 'revolutionizing,' 'innovative,' or any similar adjectives.

        Step-16: I expect you to provide a clear and consistent explanation based on the image. There's no need to mention the steps you're following, use unnecessary formatting (such as bold text), or include unrelated details, topics, or subtopics in your response. Focus solely on delivering a straightforward, cohesive explanation, without describing your process or referring to the current step. Just provide the explanation—nothing more.  
        
        Step-17: Avoid expanding abbreviations under any circumstances. Use abbreviations exactly as they appear in the extracted content. If an abbreviation is present, reproduce it as is, without repeating or expanding it at any point throughout the entire explanation.

        Step-18: If Slide 3 is detected and labeled as sketches, ignore the standard figure referencing prompt and use the following description instead:
        "Referring to Figures 3(a) and 3(b), aspects of the present disclosure include detailed sketches that illustrate two main parts: the process of light interaction with a photonic structure and a cross-sectional view of a device for sample analysis. Figure 3(a) depicts the process of light interaction with a photonic structure. In this figure, a stick figure holds a light source, directing light towards a spherical structure composed of hexagonal cells, indicating light entry into a photonic crystal. The directional flow shows light being channeled and guided through the photonic crystal, eventually exiting at specific points, possibly for analysis or detection purposes. The bottom part of Figure 3(a) shows a detailed view of the photonic crystal, where light paths are enhanced within the structure's cavity, leading to concentrated points. These points are marked as 'L' for light sources and 'M' for photodetectors, indicating locations where light is manipulated and detected. Figure 3(b) provides a cross-sectional view of a device used for sample analysis. The directional flow starts with a sample being introduced into an enhancement cavity, marked as 'L', highlighting the area where the sample interaction is maximized. Adjacent to this, 'P' indicates a position where possibly a photodetector or related component is situated to monitor the interaction. The flow continues as the sample interacts within the cavity, and Bragg gratings on either side control the entry and exit of light or other forms of energy. The flow of air is managed via 'air out' pathways, ensuring a controlled environment for the sample. Below this setup, a filter is positioned to manage the light or other flow, ensuring only desired wavelengths or particles pass through. Both Figures 3(a) and 3(b) work together to illustrate the operational flow of guiding and enhancing light through photonic structures for efficient sample analysis. Figure 3(a) explains the initial process of light manipulation through a photonic crystal, while Figure 3(b) shows the subsequent application in a device designed for analyzing samples, providing a continuous and coherent explanation of the overall process."
        """   
        data = {  
            "model": model,  
            "messages": [  
                {"role": "system", "content": "You are a helpful assistant that responds in Markdown."},  
                {"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64_image}"}}]}  
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
            insight = result["choices"][0]["message"]["content"]  
            insights.append({  
                "slide_number": slide_number,  
                "slide_title": image_data.get('slide_title', 'Untitled Slide'),  
                "insight": insight  
            })  
        else:  
            st.error(f"Error: {response.status_code} - {response.text}")  
            insights.append({  
                "slide_number": slide_number,  
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
  
def ensure_proper_spacing(text):  
    if text:  
        # Fix spacing issues after periods  
        text = re.sub(r'\.(?!\s)', '. ', text)  # Ensure space after period  
        text = re.sub(r'\s+', ' ', text)  # Ensure single space between words  
        text = re.sub(r'(\.\s+)(\w)', lambda match: match.group(1) + match.group(2).upper(), text)  # Capitalize first letter after period  
        text = text[0].upper() + text[1:]  # Ensure the first letter of the text is capitalized  
    return text  
  
  
def save_content_to_word(aggregated_content, output_file_name, extracted_images, theme):  
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
        properly_spaced_content = ensure_proper_spacing(sanitized_content)  
        doc.add_heading(f"[[{slide['slide_number']}, {sanitized_title}]]", level=1)  
        if properly_spaced_content:  # Only add content if it exists  
            doc.add_paragraph(properly_spaced_content)  
  
    # Add extracted images after the generated content  
    if extracted_images:  
        doc.add_heading("Extracted Images", level=1)  
        for idx, (image, slide_number) in enumerate(extracted_images):  
            _, buffer = cv2.imencode('.png', image)  
            image_stream = BytesIO(buffer)  
            doc.add_paragraph(f"Image from Slide {slide_number}:")  
            doc.add_picture(image_stream, width=doc.sections[0].page_width - doc.sections[0].left_margin - doc.sections[0].right_margin)  
            doc.add_paragraph("\n")  # Add space after image  
  
    # Add the theme at the end of the document  
    doc.add_heading("Overall Theme", level=1)  
    doc.add_paragraph(theme)  
  
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
  
def extract_images_from_pdf(pdf_file, top_mask, bottom_mask, left_mask, right_mask, low_quality_slides):  
    # Open the PDF file  
    pdf_document = fitz.open(pdf_file)  
    page_images = []  
  
    for page_num in range(len(pdf_document)):  
        if page_num + 1 in low_quality_slides:  
            continue  # Skip low-quality slides  
  
        page = pdf_document.load_page(page_num)  
  
        # Extract and clean the page image  
        cleaned_image = extract_and_clean_page_image(page, top_mask, bottom_mask, left_mask, right_mask)  
        if cleaned_image is not None:  
            page_images.append((cleaned_image, page_num + 1))  # Keep track of the slide number  
  
    pdf_document.close()  
    return page_images  
  
def identify_low_quality_slides(text_content, image_slides):  
    low_quality_slides = set()  
    for slide in text_content:  
        slide_number = slide['slide_number']  
        if slide_number in image_slides:  
            continue  
        word_count = len(slide['text'].split())  
        if word_count < 30:  
            low_quality_slides.add(slide_number)  
        if any(generic in slide['text'].lower() for generic in ["introduction", "thank you", "summary", "inventor details"]):  
            low_quality_slides.add(slide_number)  
    return low_quality_slides  

def upload_to_blob_storage(file_name, file_data):  
    try:  
        blob_client = blob_service_client.get_blob_client(container=container_name, blob=file_name)  
        blob_client.upload_blob(file_data, overwrite=True)  
        st.info(f"{file_name} uploaded to Azure Blob Storage.")  
    except Exception as e:  
        st.error(f"Failed to upload {file_name} to Azure Blob Storage: {e}")  
  
def download_from_blob_storage(file_name):  
    try:  
        blob_client = blob_service_client.get_blob_client(container=container_name, blob=file_name)  
        blob_data = blob_client.download_blob().readall()  
        return BytesIO(blob_data)  
    except Exception as e:  
        st.error(f"Failed to download {file_name} from Azure Blob Storage: {e}")  
        return None 
  
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
            # Save uploaded PPT to a temporary file and upload to Azure Blob Storage  
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_ppt_file:  
                temp_ppt_file.write(uploaded_ppt.read())  
                temp_ppt_file_path = temp_ppt_file.name  
              
            upload_to_blob_storage(ppt_filename, open(temp_ppt_file_path, "rb"))  
              
            # Download the PPT file from Azure Blob Storage  
            ppt_blob = download_from_blob_storage(ppt_filename)  
            if not ppt_blob:  
                st.error("Failed to download PPT file from Azure Blob Storage.")  
                return  
              
            # Convert PPT to PDF  
            with open("temp_ppt.pptx", "wb") as f:  
                f.write(ppt_blob.read())  
              
            if not ppt_to_pdf("temp_ppt.pptx", "temp_pdf.pdf"):  
                st.error("PDF conversion failed. Please check the uploaded PPT file.")  
                return  
              
            # Read the PPT file content as bytes  
            with open("temp_ppt.pptx", "rb") as f:  
                ppt_bytes = f.read()  
  
            # Extract text and identify slides with visual elements  
            text_content = extract_text_from_ppt(BytesIO(ppt_bytes))  
            visual_slides = identify_visual_elements(ppt_bytes)  
  
            # Detect slides with images  
            image_slides = detect_image_slides(ppt_bytes)  
  
            # Identify low-quality slides  
            low_quality_slides = identify_low_quality_slides(text_content, image_slides)  
  
            # Combine slide numbers from both functions  
            combined_slides = combine_slide_numbers(image_slides, visual_slides)  
  
            # Capture images of marked slides  
            slide_images = capture_slide_images("temp_pdf.pdf", combined_slides, low_quality_slides)  
  
            st.info("Generating high-level theme...")  
            theme = generate_theme(text_content, slide_images)  
  
            st.info("Generating text insights...")  
            text_insights = generate_text_insights(text_content, visual_slides, text_length, theme, low_quality_slides)  
  
            st.info("Generating image insights...")  
            image_insights = generate_image_insights(slide_images, text_length, api_key, azure_endpoint, model, api_version, theme, low_quality_slides)  

            st.info("Extracting additional images...")  
            extracted_images = extract_images_from_pdf("temp_pdf.pdf", top_mask, bottom_mask, left_mask, right_mask, low_quality_slides)  

            st.info("Aggregating content...")  
            aggregated_content = aggregate_content(text_insights, image_insights)  

            st.info("Saving to Word document...")  
            output_doc = save_content_to_word(aggregated_content, output_word_filename, extracted_images, theme)  

            # Save the final Word document to Azure Blob Storage  
            upload_to_blob_storage(output_word_filename, output_doc)  

            st.download_button(label="Download Word Document", data=output_doc, file_name=output_word_filename)  

            st.success("Processing completed successfully!")  
        except Exception as e:  
            st.error(f"An error occurred: {e}")  
  
if __name__ == "__main__":  
    main()     
