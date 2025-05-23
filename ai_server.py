from fastapi import FastAPI
from pydantic import BaseModel
from transformers import AutoTokenizer, AutoModelForCausalLM
import torch

app = FastAPI()

tokenizer = AutoTokenizer.from_pretrained("microsoft/phi-1_5")
model = AutoModelForCausalLM.from_pretrained("microsoft/phi-1_5")
device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
model.to(device)

class Input(BaseModel):
    instruction: str

@app.post("/ai-command")
def get_ai_output(data: Input):
    prompt = f"Generate Excel formula or logic for: {data.instruction}\nOutput:"
    inputs = tokenizer(prompt, return_tensors="pt").to(device)
    output = model.generate(**inputs, max_length=100)
    
    # ✅ Output cleanup inside function
    result = tokenizer.decode(output[0], skip_special_tokens=True)
    result = result.replace("Output:", "").replace("Formula:", "").strip()

    return {"output": result}
