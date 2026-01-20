# PlanTemp
建议使用python11
# 安装运行环境
pip install -r requirements.txt
# 配置api_key
## Linux/macOS
export DASHSCOPE_API_KEY="你的key"
## Windows
setx DASHSCOPE_API_KEY "你的key"或者手动设置环境变量
# 启动服务
python main.py --host 0.0.0.0 --port 8000
# 请求接口
POST "http://127.0.0.1:8000/generate"
