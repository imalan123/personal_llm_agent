import os
import re
from typing import Optional, Any

import ollama
from dotenv import load_dotenv
from langchain_core.callbacks import CallbackManagerForLLMRun
from langchain_core.language_models import LLM

load_dotenv()

class Gemma3Model(LLM):

    def _call(
            self,
            prompt: str,
            stop: Optional[list[str]] = None,
            run_manager: Optional[CallbackManagerForLLMRun] = None,
            **kwargs: Any,
    ) -> str:
        model_name = os.getenv("OLLAMA_LLM_MODEL")
        client = ollama.Client()
        response_llm = client.generate(model=model_name, prompt=prompt)
        response_only = response_llm["response"]
        cleaned_content = re.sub(r"<think>.*?</think>", "", response_only, flags=re.DOTALL)

        return cleaned_content.strip()

    def _llm_type(self) -> str:
        return "gemma3:4b"

