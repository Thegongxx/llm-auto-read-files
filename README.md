# llm-auto-read-files
llm automatically read files


## 项目结构  
Project/  
├── data/  
│   ├── raw/             # 你的旧报告（用于风格学习）  
│   └── templates/       # 空白实验报告模板  
├── db/                  # 向量数据库存储  
├── src/
│   ├── rag_engine.py    # 负责知识库检索  
│   ├── llm_client.py    # 负责调用 API  
│   ├── doc_handler.py   # 负责 Word 读写与格式设置  
│   └── main.py          # 主程序入口  
└── .env                 # API Key 配置  
 
