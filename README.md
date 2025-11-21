This project generates a PowerPoint from a text prompt using OpenAI.  
You type a prompt and the program creates slide titles, bullet points, and an optional image slide at the end.


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
python -m venv venv
.\venv\Scripts\Activate.ps1

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


Summary:

- The LLM (OpenAI) generates the slide content.
- SerpAPI searches for one related image.
- The final PPTX is fully created by Python using `python-pptx`.
