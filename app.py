import logging
import os
from pathlib import Path

from dotenv import load_dotenv

from server import create_server

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s %(name)s: %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


def main():
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8765"))

    server = create_server(host, port)

    print(f"╔══════════════════════════════════════════╗")
    print(f"║   AI 论文阅读生成 PPT                    ║")
    print(f"╠══════════════════════════════════════════╣")
    print(f"║  页面： http://{host}:{port}              ")
    print(f"║  接口： http://{host}:{port}/api/generate  ")
    print(f"╚══════════════════════════════════════════╝")
    print(f"  4 种模板风格可选 · 支持 AI / 本地双模式")
    print(f"  文件保存在 outputs/ 目录\n")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止。")


if __name__ == "__main__":
    main()
