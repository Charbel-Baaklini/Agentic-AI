This project generates a PowerPoint presentation from a text prompt using OpenAI.
You enter any topic, and the program automatically creates slide titles, bullet points, a color theme based on the topic, and images pulled from SerpAPI.

Install everything from: requirements.txt

Dependencies include:
- openai  
- python-pptx  
- python-dotenv  
- requests  
- Pillow  

Create a `.env` file in the project root:
OPENAI_API_KEY = "your_key_here"
SERPAPI_API_KEY = "your_serpapi_key_here"
CSE_ID = "your_cse_id_here"

1. Open terminal in the project folder:

2. Create and activate a venv:
    1- python -m venv venv
    2- .\venv\Scripts\Activate.ps1

3. Install dependencies:
pip install -r requirements.txt

4. Run the program:
python main.py

5.Enter the prompt for your PowerPoint slide:
Make a detailed presentation about lions, covering habitat, hunting behavior, pride structure, evolution, and threats to survival.

6.Your PowerPoint file will be created here:
outputs/output_deck.pptx

7.Any downloaded images are saved in:
outputs/images/


How It Works:
The LLM (OpenAI) generates the slide titles & bullet points.
A mini agent performs planning and verification.
A color theme is selected based on the topic.
SerpAPI searches the web for one related image.
Python builds the final .pptx using python-pptx.
