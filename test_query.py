import json
from llama_index.llms.openai import OpenAI
import requests
import datetime
from otterai import OtterAI
import os
from docx import Document as DocumentDocx
from dotenv import load_dotenv
from llama_index.core import SimpleDirectoryReader
from dotenv import load_dotenv
from llama_index.embeddings.openai import OpenAIEmbedding
from llama_index.core import Settings
from llama_index.core import VectorStoreIndex, StorageContext
from llama_index.vector_stores.milvus import MilvusVectorStore
from llama_index.core.schema import Document
from llama_index.core.service_context import ServiceContext
 

load_dotenv()


Settings.embed_model = OpenAIEmbedding(model="text-embedding-3-large", dim=3072)
Settings.context_window = 4096
Settings.llm = OpenAI(
    model='gpt-4o', temperature=0.1, max_tokens=4000, streaming=True
)

vector_store_full_transcripts = MilvusVectorStore(
    uri=os.getenv('MILVUS_SERVER'), token=os.getenv('MILVUS_API_KEY'), dim=3072, overwrite=False, collection_name=os.getenv('MILVUS_COLLECTION_NAME_FULL_TRANSCRIPTS')
)
storage_context_full_transcripts = StorageContext.from_defaults(vector_store=vector_store_full_transcripts)
index = VectorStoreIndex.from_vector_store(vector_store_full_transcripts)
query_engine = index.as_query_engine(streaming=True, similarity_top_k=20)

print (query_engine.query("Write a detailed report about everything said about security diligence for Apollo?"))