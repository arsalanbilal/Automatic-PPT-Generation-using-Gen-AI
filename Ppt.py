import streamlit as st
import pandas as pd
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import os
from langchain_groq import ChatGroq
from langchain_huggingface import HuggingFaceEmbeddings
from langchain.chains import LLMChain
from langchain.prompts import PromptTemplate
import tempfile

# Streamlit UI Configuration
st.set_page_config(
    page_title="AutoPPT Pro - AI Presentation Generator",
    page_icon="📊",
    layout="wide"
)

# Title and Description
st.title("🤖 AutoPPT Pro - AI Presentation Generator")
st.markdown("""
Generate professional PowerPoint presentations automatically using Generative AI. 
Upload your data or describe your topic, and let AI create a complete presentation for you.
""")

# API Key Section in Sidebar
st.sidebar.header("🔑 API Configuration")

# Groq API Key input
groq_api_key = st.sidebar.text_input(
    "Enter your Groq API Key:",
    type="password",
    placeholder="gsk_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
    help="Get your free API key from https://console.groq.com"
)

# Check if API key is provided
if not groq_api_key:
    st.sidebar.warning("⚠️ Please enter your Groq API key to continue")
    st.stop()

# Initialize LLM with user-provided API key
try:
    llm = ChatGroq(groq_api_key=groq_api_key, model_name="openai/gpt-oss-120b")
    # Test the API key with a simple call
    llm.invoke("Hello")  # Simple test to verify API key
    st.sidebar.success("✅ Groq API key validated!")
except Exception as e:
    st.sidebar.error(f"❌ Invalid Groq API key: {str(e)}")
    st.stop()

# Initialize Hugging Face embeddings
try:
    embeddings = HuggingFaceEmbeddings(
        model_name="sentence-transformers/all-MiniLM-L6-v2",
        model_kwargs={'device': 'cpu'}
    )
    st.sidebar.success("✅ Hugging Face embeddings loaded!")
except Exception as e:
    st.sidebar.error(f"❌ Error loading embeddings: {str(e)}")
    st.stop()

# Main Content Area
tab1, tab2, tab3 = st.tabs(["📝 Text to PPT", "📊 Data to PPT", "⚙️ Advanced Settings"])

def create_presentation_from_structure(topic, audience, slides_count, tone, additional_context):
    """Create PowerPoint presentation using AI-generated structure"""
    
    # Create presentation structure using LangChain with Groq
    ppt_prompt = PromptTemplate(
        input_variables=["topic", "audience", "slides_count", "tone", "additional_context"],
        template="""
        You are an expert presentation designer. Create a comprehensive PowerPoint presentation structure for the following:
        
        TOPIC: {topic}
        AUDIENCE: {audience}
        NUMBER OF SLIDES: {slides_count}
        TONE: {tone}
        ADDITIONAL CONTEXT: {additional_context}
        
        Please generate a structured presentation with:
        1. Title slide with topic and subtitle
        2. Agenda/Table of Contents
        3. Main content slides with clear headings and 3-5 bullet points each
        4. Summary/Conclusion slide
        5. Thank you/Q&A slide
        
        Format the response as:
        Slide 1: TITLE - [Presentation Title] | SUBTITLE - [Presentation Subtitle]
        Slide 2: AGENDA - [List main points as bullet points]
        Slide 3: HEADING - [Slide Title] | CONTENT - [Bullet point 1] | [Bullet point 2] | [Bullet point 3]
        ...
        Slide N: CONCLUSION - [Slide Title] | CONTENT - [Key takeaway 1] | [Key takeaway 2] | [Key takeaway 3]
        Slide N+1: THANK YOU - [Thank you message] | CONTACT - [Contact information if available]
        """
    )
    
    ppt_chain = LLMChain(llm=llm, prompt=ppt_prompt)
    ppt_structure = ppt_chain.run({
        "topic": topic,
        "audience": audience,
        "slides_count": slides_count,
        "tone": tone,
        "additional_context": additional_context
    })
    
    return ppt_structure

def parse_slide_structure(ppt_structure):
    """Parse the AI-generated structure into slide components"""
    slides = []
    lines = ppt_structure.split('\n')
    
    for line in lines:
        line = line.strip()
        if line.startswith('Slide') and ':' in line:
            # Extract slide content after the colon
            slide_content = line.split(':', 1)[1].strip()
            slides.append(slide_content)
    
    return slides

def create_ppt_file(slides_data, topic):
    """Create actual PowerPoint file from slide data"""
    prs = Presentation()
    
    # Title slide
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    
    # Use first slide for title if available
    if slides_data and 'TITLE' in slides_data[0]:
        title_text = slides_data[0].split('TITLE - ')[1].split(' | ')[0] if 'TITLE - ' in slides_data[0] else topic
        subtitle_text = slides_data[0].split('SUBTITLE - ')[1] if 'SUBTITLE - ' in slides_data[0] else "AI-Generated Presentation"
    else:
        title_text = topic
        subtitle_text = "AI-Generated Presentation"
    
    title.text = title_text
    subtitle.text = subtitle_text
    
    # Add content slides (skip first slide if it was title)
    start_index = 1 if slides_data and 'TITLE' in slides_data[0] else 0
    
    for i in range(start_index, len(slides_data)):
        slide_content = slides_data[i]
        
        # Use bullet slide layout
        bullet_slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(bullet_slide_layout)
        title_shape = slide.shapes.title
        body_shape = slide.placeholders[1]
        
        # Extract title and content
        if ' - ' in slide_content:
            parts = slide_content.split(' - ', 1)
            slide_title = parts[0].replace('HEADING:', '').replace('AGENDA:', '').replace('CONCLUSION:', '').replace('THANK YOU:', '').strip()
            
            title_shape.text = slide_title
            
            # Add content to body
            if 'CONTENT - ' in parts[1]:
                content_text = parts[1].split('CONTENT - ')[1]
                bullet_points = [point.strip() for point in content_text.split('|')]
            else:
                bullet_points = [point.strip() for point in parts[1].split('|')]
            
            tf = body_shape.text_frame
            tf.text = bullet_points[0] if bullet_points else "Key points"
            
            for point in bullet_points[1:]:
                if point and point != "CONTENT":
                    p = tf.add_paragraph()
                    p.text = point
        else:
            title_shape.text = f"Slide {i+1}"
            body_shape.text_frame.text = slide_content
    
    # Save to bytes
    ppt_bytes = io.BytesIO()
    prs.save(ppt_bytes)
    ppt_bytes.seek(0)
    
    return ppt_bytes

with tab1:
    st.header("Generate PPT from Text Description")
    
    topic = st.text_input("Presentation Topic:", placeholder="e.g., Digital Transformation Strategy 2024")
    audience = st.selectbox("Target Audience:", ["Executive Leadership", "Technical Team", "Sales & Marketing", "General Audience", "Students"])
    slides_count = st.slider("Number of Slides:", 5, 20, 10)
    tone = st.selectbox("Presentation Tone:", ["Professional", "Persuasive", "Educational", "Inspirational", "Technical"])
    
    additional_context = st.text_area("Additional Context (Optional):", 
                                    placeholder="Any specific points to include, company information, or special requirements...")
    
    if st.button("Generate Presentation", key="text_to_ppt"):
        if not topic:
            st.error("Please enter a presentation topic!")
        else:
            with st.spinner("🤖 AI is generating your presentation..."):
                try:
                    # Generate presentation structure
                    ppt_structure = create_presentation_from_structure(
                        topic, audience, slides_count, tone, additional_context
                    )
                    
                    # Parse the structure
                    slides_data = parse_slide_structure(ppt_structure)
                    
                    # Create PowerPoint file
                    ppt_bytes = create_ppt_file(slides_data, topic)
                    
                    # Download button
                    st.success("🎉 Presentation generated successfully!")
                    
                    col1, col2 = st.columns([1, 2])
                    
                    with col1:
                        st.download_button(
                            label="📥 Download PowerPoint",
                            data=ppt_bytes,
                            file_name=f"{topic.replace(' ', '_')}_presentation.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                    
                    with col2:
                        st.info(f"**Generated {len(slides_data)} slides** using Groq LLM and Hugging Face embeddings")
                    
                    # Show generated content
                    with st.expander("📋 View AI-Generated Content"):
                        st.write("### Presentation Structure")
                        st.text_area("Raw Structure:", ppt_structure, height=200)
                        
                        st.write("### Parsed Slides")
                        for i, slide in enumerate(slides_data):
                            st.write(f"**Slide {i+1}:** {slide}")
                            st.write("---")
                            
                except Exception as e:
                    st.error(f"Error generating presentation: {str(e)}")

with tab2:
    st.header("Generate PPT from Data")
    
    uploaded_file = st.file_uploader("Upload CSV/Excel file", type=['csv', 'xlsx'])
    
    if uploaded_file:
        # Save uploaded file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix=".csv" if uploaded_file.name.endswith('.csv') else ".xlsx") as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(tmp_file_path)
            else:
                df = pd.read_excel(tmp_file_path)
            
            st.write("### Data Preview")
            st.dataframe(df.head())
            
            # Basic data analysis
            st.write("### Data Summary")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Total Rows", len(df))
            with col2:
                st.metric("Total Columns", len(df.columns))
            with col3:
                st.metric("Data Types", f"{len(df.select_dtypes(include='number').columns)} numeric, {len(df.select_dtypes(include='object').columns)} text")
            
            analysis_type = st.selectbox("Analysis Type:", 
                                       ["Data Summary", "Trend Analysis", "Comparative Analysis", "Key Insights"])
            
            if st.button("Generate Data Presentation", key="data_to_ppt"):
                with st.spinner("🤖 Analyzing data and creating presentation..."):
                    try:
                        # Create data analysis prompt
                        data_prompt = PromptTemplate(
                            input_variables=["data_preview", "analysis_type", "columns", "row_count"],
                            template="""
                            Analyze the following dataset and create a presentation structure for {analysis_type}:
                            
                            DATA PREVIEW:
                            {data_preview}
                            
                            DATASET INFO:
                            - Columns: {columns}
                            - Total Rows: {row_count}
                            
                            Create a presentation structure with:
                            1. Title slide about the data analysis
                            2. Dataset overview and methodology
                            3. Key findings and insights
                            4. Visualizations recommendations
                            5. Conclusions and recommendations
                            
                            Format the response as:
                            Slide 1: TITLE - [Analysis Title] | SUBTITLE - [Dataset Description]
                            Slide 2: OVERVIEW - [Slide Title] | CONTENT - [Key point 1] | [Key point 2] | [Key point 3]
                            ...
                            """
                        )
                        
                        data_chain = LLMChain(llm=llm, prompt=data_prompt)
                        data_structure = data_chain.run({
                            "data_preview": df.head().to_string(),
                            "analysis_type": analysis_type,
                            "columns": ", ".join(df.columns.tolist()),
                            "row_count": len(df)
                        })
                        
                        slides_data = parse_slide_structure(data_structure)
                        ppt_bytes = create_ppt_file(slides_data, f"Data Analysis - {analysis_type}")
                        
                        st.success("🎉 Data presentation generated!")
                        st.download_button(
                            label="📥 Download Data Presentation",
                            data=ppt_bytes,
                            file_name=f"data_analysis_{analysis_type.replace(' ', '_')}.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                        
                        with st.expander("📊 View Analysis Results"):
                            st.text(data_structure)
                            
                    except Exception as e:
                        st.error(f"Error in data analysis: {str(e)}")
        
        except Exception as e:
            st.error(f"Error reading file: {str(e)}")
        finally:
            # Clean up temporary file
            import os
            if os.path.exists(tmp_file_path):
                os.unlink(tmp_file_path)

with tab3:
    st.header("Advanced Settings")
    
    st.subheader("Presentation Template")
    template_choice = st.selectbox("Choose Template:", 
                                 ["Corporate Blue", "Modern Red", "Professional Green", "Creative Purple"])
    
    st.subheader("AI Model Settings")
    st.info(f"Current Model: Groq - openai/gpt-oss-120b")
    st.info(f"Embeddings: Hugging Face - sentence-transformers/all-MiniLM-L6-v2")
    
    st.subheader("Content Options")
    col1, col2 = st.columns(2)
    
    with col1:
        include_agenda = st.checkbox("Include Agenda Slide", value=True)
        include_summary = st.checkbox("Include Summary Slide", value=True)
        include_qa = st.checkbox("Include Q&A Slide", value=True)
    
    with col2:
        add_speaker_notes = st.checkbox("Add Speaker Notes", value=False)
        add_references = st.checkbox("Add References Slide", value=False)
        detailed_bullets = st.checkbox("Detailed Bullet Points", value=True)

# Instructions Section
with st.expander("📖 How to Use This Tool"):
    st.markdown("""
    ### Step-by-Step Guide:
    
    1. **Configure API**: Enter your Groq API key in the sidebar
    2. **Choose Input Method**:
       - **Text to PPT**: Describe your topic and requirements
       - **Data to PPT**: Upload your dataset for analysis-based presentations
    3. **Customize Settings**: Adjust audience, tone, and slide count
    4. **Generate**: Click the generate button and wait for AI to create your presentation
    5. **Download**: Get your professionally formatted PowerPoint file
    
    ### Technical Stack:
    - **LLM**: Groq (openai/gpt-oss-120b) for content generation
    - **Embeddings**: Hugging Face (sentence-transformers/all-MiniLM-L6-v2)
    - **Framework**: LangChain for orchestration
    - **UI**: Streamlit for web interface
    - **PPT Generation**: python-pptx library
    
    ### Features:
    - AI-powered content generation using Groq LLM
    - Professional PowerPoint formatting
    - Data analysis and visualization recommendations
    - Multiple template options
    - Enterprise-ready presentation output
    """)

# Footer
st.markdown("---")
st.markdown(
    "🔒 Built with Groq LLM & Hugging Face Embeddings | " +
    "Powered by LangChain & Streamlit",
    unsafe_allow_html=True
)
