# -*- coding: utf-8 -*-
"""
股票数据处理总控脚本
功能：
1. 选择市场（US/A/HK）
2. 清除对应 demo 目录下的数据文件
3. 按顺序执行各处理脚本
4. 任一步骤失败则暂停后续执行
5. 通过钉钉/飞书 Webhook 发送通知
6. 运行日志写入日志文件
"""

import os
import sys
import shutil
import subprocess
import time
import json
import hmac
import hashlib
import base64
import urllib.parse
import logging
import requests
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv

# 加载 .env 配置
load_dotenv()

# ===================== 日志配置 =====================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_NAME = Path(__file__).stem  # run_pipeline
LOG_TIME = datetime.now().strftime("%Y%m%d_%H%M%S")
LOG_FILE = os.path.join(SCRIPT_DIR, f"{SCRIPT_NAME}_log_{LOG_TIME}.txt")

# 配置日志：同时输出到控制台和文件
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ]
)
logger = logging.getLogger(__name__)

# ===================== 配置区 =====================
BASE_DIR = r"/home/ubuntu/alibaba_agent/downlaods_futu_data/data_all_market"

# 处理流程（按顺序执行）
PIPELINE = [
    {"name": "FROM_CSV_XLSX_TO_DEMO1-1.py", "desc": "获取富途行情数据"},
    {"name": "ReadExcelFileCaculate1-2yun.py", "desc": "计算价格区间"},
    {"name": "get_double16splitagain1-3yun.py", "desc": "双区间匹配处理"},
    {"name": "get2okx标黄可行1-4okx.py", "desc": "标黄处理"},
    {"name": "get31process_mark.py", "desc": "最终标注处理"},
    {"name": "get32process_精减.py", "desc": "精减标注数据"},
]

# 需要清除的文件扩展名
CLEAR_EXTENSIONS = {".xlsx", ".xls", ".csv", ".txt"}
# ============================================================


def get_config():
    """读取通知配置"""
    return {
        "notify_mode": os.getenv("NOTIFY_MODE", "dingtalk"),
        "dingtalk_webhook": os.getenv("DINGTALK_WEBHOOK", ""),
        "dingtalk_secret": os.getenv("DINGTALK_SECRET", ""),
        "feishu_webhook": os.getenv("FEISHU_WEBHOOK", ""),
        "feishu_secret": os.getenv("FEISHU_SECRET", ""),
    }


def format_duration(seconds):
    """格式化耗时"""
    if seconds < 60:
        return f"{seconds:.1f} 秒"
    elif seconds < 3600:
        minutes = int(seconds // 60)
        secs = int(seconds % 60)
        return f"{minutes} 分 {secs} 秒"
    else:
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        return f"{hours} 小时 {minutes} 分"


def generate_dingtalk_sign(secret):
    """生成钉钉加签"""
    if not secret:
        return ""
    timestamp = str(round(time.time() * 1000))
    string_to_sign = f"{timestamp}\n{secret}"
    hmac_code = hmac.new(
        secret.encode("utf-8"),
        string_to_sign.encode("utf-8"),
        digestmod=hashlib.sha256
    ).digest()
    sign = urllib.parse.quote_plus(base64.b64encode(hmac_code))
    return f"&timestamp={timestamp}&sign={sign}"


def send_dingtalk(webhook, secret, title, content):
    """发送钉钉通知"""
    if not webhook:
        print("⚠️ 钉钉 Webhook 未配置，跳过")
        return False
    
    url = webhook
    if secret:
        url += generate_dingtalk_sign(secret)
    
    data = {
        "msgtype": "markdown",
        "markdown": {
            "title": title,
            "text": content
        }
    }
    
    try:
        resp = requests.post(url, json=data, timeout=10)
        result = resp.json()
        if result.get("errcode") == 0:
            print("✅ 钉钉通知发送成功")
            return True
        else:
            print(f"❌ 钉钉通知发送失败: {result}")
            return False
    except Exception as e:
        print(f"❌ 钉钉通知异常: {e}")
        return False


def generate_feishu_sign(secret):
    """生成飞书加签"""
    if not secret:
        return "", ""
    timestamp = str(int(time.time()))
    string_to_sign = f"{timestamp}\n{secret}"
    hmac_code = hmac.new(
        string_to_sign.encode("utf-8"),
        digestmod=hashlib.sha256
    ).digest()
    sign = base64.b64encode(hmac_code).decode("utf-8")
    return timestamp, sign


def send_feishu(webhook, secret, title, content, is_success=True):
    """发送飞书通知（加签方式）"""
    if not webhook:
        print("⚠️ 飞书 Webhook 未配置，跳过")
        return False
    
    # 生成加签
    timestamp, sign = generate_feishu_sign(secret)
    
    # 根据状态设置颜色
    header_color = "green" if is_success else "red"
    
    # 将 Markdown 转换为飞书格式
    data = {
        "msg_type": "interactive",
        "timestamp": timestamp,
        "sign": sign,
        "card": {
            "header": {
                "title": {
                    "tag": "plain_text",
                    "content": title
                },
                "template": header_color
            },
            "elements": [
                {
                    "tag": "markdown",
                    "content": content
                }
            ]
        }
    }
    
    try:
        resp = requests.post(webhook, json=data, timeout=10)
        result = resp.json()
        if result.get("code") == 0:
            print("✅ 飞书通知发送成功")
            return True
        else:
            print(f"❌ 飞书通知发送失败: {result}")
            return False
    except Exception as e:
        print(f"❌ 飞书通知异常: {e}")
        return False


def send_notification(title, content, is_success=True):
    """根据配置发送通知"""
    config = get_config()
    mode = config["notify_mode"].lower()
    
    results = []
    
    if mode in ("dingtalk", "both"):
        ok = send_dingtalk(
            config["dingtalk_webhook"],
            config["dingtalk_secret"],
            title,
            content
        )
        results.append(("钉钉", ok))
    
    if mode in ("feishu", "both"):
        ok = send_feishu(
            config["feishu_webhook"],
            config["feishu_secret"],
            title,
            content,
            is_success
        )
        results.append(("飞书", ok))
    
    return results


def clear_demo_directory(market):
    """清除 demo 目录及子目录下的数据文件"""
    demo_dir = os.path.join(BASE_DIR, market, "demo")
    
    if not os.path.exists(demo_dir):
        print(f"📂 目录不存在，跳过清除: {demo_dir}")
        return 0
    
    cleared_count = 0
    
    for root, dirs, files in os.walk(demo_dir):
        for file in files:
            ext = os.path.splitext(file)[1].lower()
            if ext in CLEAR_EXTENSIONS:
                file_path = os.path.join(root, file)
                try:
                    os.remove(file_path)
                    cleared_count += 1
                    print(f"  🗑️ 已删除: {file}")
                except Exception as e:
                    print(f"  ❌ 删除失败 {file}: {e}")
    
    print(f"\n✅ 共清除 {cleared_count} 个数据文件")
    return cleared_count


def run_script(script_name, market):
    """运行单个脚本"""
    script_path = os.path.join(SCRIPT_DIR, script_name)
    
    if not os.path.exists(script_path):
        logger.error(f"脚本不存在: {script_path}")
        return False, f"脚本不存在: {script_path}"
    
    # 临时修改脚本中的 MARKET 变量
    # 通过环境变量传递市场选择
    env = os.environ.copy()
    env["MARKET"] = market

    # 为精减脚本设置输入目录
    if script_name == "get32process_精减.py":
        process_input_dir = os.path.join(
            BASE_DIR, market, "demo", "计算区间", "双区间相同再算幅度", "标黄", "已标注"
        )
        env["PROCESS_INPUT_DIR"] = process_input_dir

    try:
        logger.info(f"{'='*60}")
        logger.info(f"▶️ 正在运行: {script_name}")
        logger.info(f"{'='*60}")
        
        result = subprocess.run(
            [sys.executable, script_path],
            env=env,
            capture_output=True,
            text=True,
            timeout=3600  # 1小时超时
        )
        
        if result.stdout:
            logger.info(result.stdout)
        
        if result.returncode != 0:
            logger.error(f"脚本执行失败，返回码: {result.returncode}")
            if result.stderr:
                logger.error(f"错误信息: {result.stderr}")
            return False, result.stderr or result.stdout
        
        logger.info(f"✅ {script_name} 执行成功")
        return True, result.stdout
        
    except subprocess.TimeoutExpired:
        logger.error(f"脚本执行超时（超过1小时）: {script_name}")
        return False, "脚本执行超时（超过1小时）"
    except Exception as e:
        logger.error(f"脚本执行异常: {script_name}, 错误: {e}")
        return False, str(e)


def select_market():
    """手动选择市场"""
    print("\n" + "="*40)
    print("请选择要处理的市场:")
    print("  1. US (美股)")
    print("  2. A  (A股)")
    print("  3. HK (港股)")
    print("="*40)
    
    while True:
        choice = input("\n请输入选项 (1/2/3 或 US/A/HK): ").strip().upper()
        
        if choice in ("1", "US"):
            return "US"
        elif choice in ("2", "A"):
            return "A"
        elif choice in ("3", "HK"):
            return "HK"
        else:
            print("❌ 无效选项，请重新输入")


def get_markets_by_mode():
    """根据配置获取要运行的市场列表"""
    mode = os.getenv("MARKET_MODE", "MANUAL").strip().upper()
    
    if mode == "US":
        return ["US"]
    elif mode == "A":
        return ["A"]
    elif mode == "HK":
        return ["HK"]
    elif mode == "AUTO":
        return ["US", "A", "HK"]
    elif mode == "MANUAL":
        return [select_market()]
    else:
        logger.warning(f"未知的 MARKET_MODE: {mode}，使用手动模式")
        return [select_market()]


def run_market_pipeline(market, auto_confirm=False):
    """运行单个市场的处理流程，返回 (is_success, results_dict)"""
    start_time = time.time()
    start_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    logger.info(f"🎯 开始处理市场: {market}")
    logger.info(f"📂 基础目录: {BASE_DIR}/{market}")
    
    # 确认清除（非自动模式才询问）
    if not auto_confirm:
        confirm = input(f"\n⚠️ 即将清除 {market}/demo 目录下的所有数据文件，是否继续? (y/n): ").strip().lower()
        if confirm != "y":
            logger.warning(f"用户取消 {market} 市场操作")
            return False, {"market": market, "status": "已取消", "results": [], "cleared_count": 0}
    
    # 清除目录
    logger.info(f"🧹 正在清除 {market}/demo 目录...")
    cleared_count = clear_demo_directory(market)
    
    # 执行处理流程
    results = []
    success_count = 0
    failed_step = None
    
    for step in PIPELINE:
        script_name = step["name"]
        desc = step["desc"]
        
        logger.info(f"📋 步骤: {desc}")
        success, output = run_script(script_name, market)
        
        results.append({
            "name": script_name,
            "desc": desc,
            "success": success,
            "output": output[:500] if output else ""
        })
        
        if success:
            success_count += 1
        else:
            failed_step = step
            logger.error(f"❌ 步骤失败: {desc}")
            logger.error("⏸️ 后续步骤已暂停")
            break
    
    # 统计结果
    end_time = time.time()
    duration = end_time - start_time
    end_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    is_success = failed_step is None
    status = "✅ 全部成功" if is_success else "❌ 部分失败"
    
    result_dict = {
        "market": market,
        "status": status,
        "is_success": is_success,
        "start_time": start_time_str,
        "end_time": end_time_str,
        "duration": duration,
        "results": results,
        "success_count": success_count,
        "total_steps": len(PIPELINE),
        "cleared_count": cleared_count,
        "failed_step": failed_step,
    }
    
    logger.info(f"{'='*60}")
    logger.info(f"市场 {market} 处理完成!")
    logger.info(f"  状态: {status}")
    logger.info(f"  成功: {success_count}/{len(PIPELINE)}")
    logger.info(f"  耗时: {format_duration(duration)}")
    logger.info(f"{'='*60}")
    
    return is_success, result_dict


def build_notification_message(market, result_dict):
    """构建单个市场的通知消息"""
    status = result_dict["status"]
    title = f"股票数据处理 - {market}市场 - {status}"
    
    message = f"""## 股票数据处理报告

**市场**: {market}
**状态**: {status}
**开始时间**: {result_dict['start_time']}
**结束时间**: {result_dict['end_time']}
**总耗时**: {format_duration(result_dict['duration'])}

---

### 执行步骤

"""
    
    for i, r in enumerate(result_dict["results"], 1):
        status_icon = "✅" if r["success"] else "❌"
        message += f"{i}. {status_icon} **{r['desc']}**\n"
    
    message += f"""
---

### 统计信息

- 成功: {result_dict['success_count']}/{result_dict['total_steps']}
- 清除文件: {result_dict['cleared_count']} 个
- 输出目录: `{BASE_DIR}/{market}/demo`
"""
    
    if result_dict["failed_step"]:
        message += f"\n> **失败步骤**: {result_dict['failed_step']['desc']}"
    
    return title, message


def build_summary_notification(all_results):
    """构建 AUTO 模式的汇总通知消息"""
    total_markets = len(all_results)
    success_markets = sum(1 for r in all_results if r.get("is_success"))
    all_success = success_markets == total_markets
    status = "✅ 全部成功" if all_success else "❌ 部分失败"
    
    title = f"股票数据处理(全部市场) - {status}"
    
    message = f"""## 股票数据处理报告（全部市场）

**运行模式**: AUTO
**状态**: {status}
**市场数**: {total_markets} 个（成功 {success_markets} 个）

---

### 各市场结果

"""
    
    for r in all_results:
        market = r["market"]
        market_status = r["status"]
        message += f"- {market_status} **{market}** ({r['success_count']}/{r['total_steps']})\n"
    
    message += f"""
---

### 详细信息

"""
    
    for r in all_results:
        market = r["market"]
        message += f"\n#### {market} 市场\n"
        message += f"- 耗时: {format_duration(r['duration'])}\n"
        for i, step in enumerate(r["results"], 1):
            icon = "✅" if step["success"] else "❌"
            message += f"  {i}. {icon} {step['desc']}\n"
        if r["failed_step"]:
            message += f"\n> 失败: {r['failed_step']['desc']}\n"
    
    return title, message


def main():
    logger.info(f"日志文件: {LOG_FILE}")
    logger.info(f"开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 获取市场模式
    market_mode = os.getenv("MARKET_MODE", "MANUAL").strip().upper()
    logger.info(f"📋 市场模式: {market_mode}")
    
    # 获取要运行的市场列表
    markets = get_markets_by_mode()
    logger.info(f"🎯 待处理市场: {', '.join(markets)}")
    
    # AUTO 模式下自动确认，其他模式手动确认
    auto_confirm = (market_mode == "AUTO")
    
    all_results = []
    
    for market in markets:
        is_success, result_dict = run_market_pipeline(market, auto_confirm=auto_confirm)
        all_results.append(result_dict)
        
        # 发送单个市场的通知
        if result_dict.get("status") != "已取消":
            title, message = build_notification_message(market, result_dict)
            logger.info("📤 正在发送通知...")
            send_notification(title, message, result_dict.get("is_success", False))
    
    # AUTO 模式下额外发送汇总通知
    if market_mode == "AUTO" and len(markets) > 1:
        title, message = build_summary_notification(all_results)
        logger.info("📤 正在发送汇总通知...")
        all_success = all(r.get("is_success") for r in all_results)
        send_notification(title, message, all_success)
    
    # 最终输出
    logger.info(f"\n{'='*60}")
    logger.info("所有市场处理完成!")
    for r in all_results:
        logger.info(f"  {r['market']}: {r['status']}")
    logger.info(f"{'='*60}")


if __name__ == "__main__":
    main()
