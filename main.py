import sys
import asyncio
from pathlib import Path
from datetime import datetime

# Wir brauchen hier nur 'pipeline' und 'Config'-relevante Utils
import pipeline
from utils import get_and_create_next_run_dir, RESET

class Config:
    PPTX_INPUT = Path('./input/Algorithmik.pptx')
    BASE_OUTPUT_PATH = Path("./output")
    RULES_FILE = 'converting_rules.yaml'
    TEX_FILENAME = "document"
    TIMESTAMP = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    
    SKIP_EXTRACTION = False 
    
    EXISTING_JSON_PATH = Path("./output/2025-12-04_12-06-51/Algorithmik_cleaned.json") 
    # ---------------------

    BASE_RESULTS_DIR = Path("Results")
    RESULTS_DIR = get_and_create_next_run_dir(BASE_RESULTS_DIR)
    
    OUTPUT_DIR = BASE_OUTPUT_PATH / TIMESTAMP
    MEDIA_OUTPUT_DIR = RESULTS_DIR / 'extracted_media'
    
    # Input Logic
    if SKIP_EXTRACTION:
        RAW_JSON_INPUT = EXISTING_JSON_PATH
    else:
        RAW_JSON_INPUT = OUTPUT_DIR / (PPTX_INPUT.stem + '.json')

    CLEANED_JSON_OUTPUT = OUTPUT_DIR / (PPTX_INPUT.stem + "_cleaned.json")
    AGENT_MAX_RETRIES = 3    
    #AGENT_LLM_MODEL = 'deepseek-coder:6.7b-instruct' very bad
    AGENT_LLM_MODEL = 'qwen3:8b'           
    #AGENT_LLM_MODEL = 'qwen2.5-coder:14b'   Long wait, bad results

    @classmethod
    def setup_directories(cls):
        cls.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        cls.MEDIA_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


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
            print(f"\n{pipeline.GREEN}SUCCESS: Pipeline finished successfully.{RESET}")
        else:
            print(f"\n{pipeline.YELLOW}FINISHED: Pipeline finished with compilation errors.{RESET}")

    except FileNotFoundError as e:
        print(f"\n{pipeline.RED}CRITICAL ERROR: File missing.{RESET}")
        print(f"{pipeline.RED}{e}{RESET}")
        sys.exit(1)
    except Exception as e:
        print(f"\n{pipeline.RED}CRITICAL ERROR: An unexpected error occurred.{RESET}")
        print(f"{pipeline.RED}{e}{RESET}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    asyncio.run(run_pipeline())