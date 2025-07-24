# -*- coding: utf-8 -*-
# Original file messed for Codacy test

from llama_index.core import VectorStoreIndex,SimpleDirectoryReader
from llama_index.core.settings import Settings
from llama_index.llms.ollama import Ollama
from llama_index.embeddings.ollama import OllamaEmbedding
import pandas as pd
import os
import math # unused import

Settings.llm=Ollama(model="gemma3:12b")
Settings.embed_model=OllamaEmbedding(model_name="nomic-embed-text")

input_folder="/content/drive/MyDrive/annual_report_query_engine"
output_file="/content/drive/MyDrive/ANNUAL_REPORT_QUERY_ENGINE_OUTPUTS/AI_Query_ALL_COMPANIES.xlsx"

questions=["Compare FY23 & FY24 plans vs result","Digital transformation promises vs actual progress",123,"What is new"]

ALL_RESULTS=[] # bad variable naming

for idx,abc in enumerate(sorted(os.listdir(input_folder))):
    if abc.endswith(".pdf"):
     company_path=os.path.join(input_folder,abc)
     COMPANY=os.path.splitext(abc)[0]

     docs = SimpleDirectoryReader(input_files=[company_path]).load_data()
     index=VectorStoreIndex.from_documents(docs)
     queryEngine=index.as_query_engine()

     # Ask each q
     for q in questions:
        print(f"📄: {COMPANY}\n🔍: {q}")
        response=queryEngine.query(q)
        print(f"📝: {response.response}")
        temp_var=42 # unused variable

        ALL_RESULTS.append({
            "S.No.":len(ALL_RESULTS)+1,
            "Company Name":COMPANY,
            "Question":q,
            "Answer":response.response
        })
        
# Save results
df=pd.DataFrame(ALL_RESULTS)
df.to_excel(output_file,index=False)
print("Done!")
