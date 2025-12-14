import sys
import asyncio
from pathlib import Path
from datetime import datetime
import pipeline
from utils import get_and_create_next_run_dir, RED,GREEN,YELLOW,RESET

class Config:
    PPTX_INPUT = Path('./input/Algorithmik.pptx')
    RULES_FILE = 'converting_rules.yaml'
    TEX_FILENAME = "document"
    
    TIMESTAMP = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    
    SKIP_EXTRACTION = False 
    
    EXISTING_JSON_PATH = Path("./output/2025-12-04_12-06-51/Algorithmik_cleaned.json") 
    
    # ---------------------
    # DIRECTORY STRUCTURE
    # ---------------------
    BASE_RESULTS_DIR = Path("Results")
    RESULTS_DIR = get_and_create_next_run_dir(BASE_RESULTS_DIR)
    OUTPUT_DIR = RESULTS_DIR 
    MEDIA_OUTPUT_DIR = RESULTS_DIR / 'extracted_media'
    JSON_OUTPUT_DIR = RESULTS_DIR / 'json_data'

    if SKIP_EXTRACTION:
        # Lese vom alten Pfad
        RAW_JSON_INPUT = EXISTING_JSON_PATH
    else:
        # Speichere das neue rohe JSON in den neuen JSON-Ordner
        RAW_JSON_INPUT = JSON_OUTPUT_DIR / (PPTX_INPUT.stem + '.json')

    # Das Cleaned JSON kommt IMMER in den aktuellen Run-Ordner (JSON Subfolder)
    CLEANED_JSON_OUTPUT = JSON_OUTPUT_DIR / (PPTX_INPUT.stem + "_cleaned.json")
    

    AGENT_MAX_RETRIES = 3    
    AGENT_LLM_MODEL = 'qwen3:8b' 

    @classmethod
    def setup_directories(cls):
        """Erstellt alle notwendigen Ordner"""
        cls.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        cls.MEDIA_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        cls.JSON_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


async def run_pipeline():
    Config.setup_directories()

    try:
        # Step 0: Extraction (Optional)
        if not Config.SKIP_EXTRACTION:
            await pipeline.step_extract_structure(Config)
        else:
            print(f"Skipping PPTX extraction. Using existing JSON: {Config.RAW_JSON_INPUT}")


        pipeline.step_extract_media(Config)
        pipeline.step_process_and_optimize_data(Config)
        latex_code = pipeline.step_generate_latex(Config)
        success = pipeline.step_save_and_compile(Config, latex_code)

        if success:
            print(f"\n{GREEN}SUCCESS: Pipeline finished successfully.{RESET}")
        else:
            print(f"\n{YELLOW}FINISHED: Pipeline finished with compilation errors.{RESET}")

    except FileNotFoundError as e:
        print(f"\n{RED}CRITICAL ERROR: File missing.{RESET}")
        print(f"{RED}{e}{RESET}")
        sys.exit(1)
    except Exception as e:
        print(f"\n{RED}CRITICAL ERROR: An unexpected error occurred.{RESET}")
        print(f"{RED}{e}{RESET}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    asyncio.run(run_pipeline())