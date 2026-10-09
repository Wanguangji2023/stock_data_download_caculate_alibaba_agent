# -*- coding: utf-8 -*-
"""
统一版：根据输入文件类型（CSV / XLSX）自动选择读取方式
=====================================================
合并自 To_excelfile1-1yunCSV.py（CSV输入）和 To_excelfile1-2yun_xlsx.py（XLSX输入）
  - 自动根据 readfile 后缀判断输入格式
  - CSV  -> load_stock_codes()（优先 utilnew_yun，失败回退自动编码读取）
  - XLSX -> get_stock_list_xlsx()（读取"筛选结果"工作表 A 列）
  - 其余逻辑（黑名单、断点续传、批量获取、降级）完全保留
"""
from futu import *
import pandas as pd
import os
import time
import sys

# ===== 关键：设置SDK底层线程为守护线程，防止IO线程僵死 =====
SysConfig.set_all_thread_daemon(True)

# ===================== 配置区（只改这里） =====================
# 市场选择：US / A / HK（改这一个变量即可切换市场，或通过环境变量 MARKET 传入）
MARKET = os.environ.get("MARKET", "US")

# 基础路径（自动根据 MARKET 生成输入/输出目录）
BASE_DIR = r"/home/ubuntu/alibaba_agent/downlaods_futu_data/data_all_market"#/home/ubuntu/alibaba_agent/downlaods_futu_data/data_all_market/US/data
FilePath = os.path.join(BASE_DIR, MARKET, "data")        # 输入目录
OUTPUT_DIR = os.path.join(BASE_DIR, MARKET, "demo")      # 输出目录
# 程序会自动扫描输入目录下的所有 CSV/XLSX 文件并逐个处理

HOST = "127.0.0.1"
PORT = 11111

SLEEP_INTERVAL = 1.5          # 正常请求间隔
RETRY_SLEEP = 3.0             # 通信出错后的休眠
MAX_RETRY = 2                 # 单次请求最大重试次数
AUTO_SAVE_INTERVAL = 50       # 每获取多少条自动保存
BLACK_LIST_FILENAME = "invalid_stock_code.txt"
BATCH_SIZE = 50               # 每批股票数量（可调小以降低整批失败概率）
BLACK_KEYWORDS = ["未知股票", "暂不提供", "otc"]
# ============================================================

cols = ["股票代码", "最新价格", "最新价格区间", "历史最低价格", "历史最低区间", "历史最高价格"]


# ============================================================
# 一、文件读取（自动判断 CSV / XLSX）
# ============================================================
def detect_encoding(path):
    """检测文件编码，优先 chardet，失败回退 utf-8"""
    try:
        import chardet
        with open(path, "rb") as f:
            raw = f.read(1024 * 1024)
        return chardet.detect(raw).get("encoding") or "utf-8"
    except Exception:
        return "utf-8"


def read_codes_auto(csv_path):
    """内置自动编码读取 CSV：按 检测编码->utf-8-sig->gb18030->latin-1 逐个尝试"""
    enc = detect_encoding(csv_path)
    df = None
    for e in dict.fromkeys([enc, "utf-8-sig", "gb18030", "latin-1"]):
        try:
            df = pd.read_csv(csv_path, encoding=e)
            break
        except Exception:
            continue
    if df is None:
        print(f"❌ 无法解码 CSV：{csv_path}，请检查文件编码")
        sys.exit(1)
    col = None
    for c in df.columns:
        if str(c).lower() in ("code", "codes", "股票代码", "代码", "证券代码"):
            col = c
            break
    if col is None:
        col = df.columns[0]
    codes = [str(x).strip() for x in df[col].tolist()]
    return codes


def load_stock_codes(csv_path):
    """读取 CSV 股票代码列表：优先用 utilnew_yun（保留A股ST过滤等逻辑），失败回退内置读取"""
    if not os.path.exists(csv_path):
        print(f"❌ CSV 文件不存在：{csv_path}")
        sys.exit(1)
    try:
        import utilnew_yun as un
        f = os.path.basename(csv_path).upper()
        if "US" in f:
            list_us = un.readCSV_US(csv_path)
            print(list_us)
            return list_us
        elif "HK" in f:
            list_hk = un.readCSV_HK(csv_path)
            print(list_hk)
            return list_hk
        else:
            list_a = un.readCSV_A(csv_path)
            print(list_a)
            return list_a
    except Exception as e:
        print(f"⚠️ 用 utilnew_yun 读取失败（{e}），回退内置自动编码读取...")
        return read_codes_auto(csv_path)


def get_stock_list_xlsx(xlsx_path):
    """读取 XLSX 股票代码列表：从"筛选结果"工作表 A 列读取"""
    sheet_name = "筛选结果"
    df = pd.read_excel(xlsx_path, sheet_name=sheet_name, usecols="A")
    stock_list = df.iloc[:, 0].dropna().astype(str).tolist()
    return stock_list


def load_stock_codes_by_type(file_path):
    """根据文件后缀自动选择读取方式"""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".csv":
        print("📄 检测到 CSV 输入，使用 CSV 读取方式")
        return load_stock_codes(file_path)
    elif ext in (".xlsx", ".xls"):
        print("📄 检测到 XLSX 输入，使用 XLSX 读取方式")
        return get_stock_list_xlsx(file_path)
    else:
        print(f"❌ 不支持的文件格式：{ext}，仅支持 .csv / .xlsx / .xls")
        sys.exit(1)


# ============================================================
# 二、黑名单管理
# ============================================================
def get_black_list():
    """读取黑名单文件，返回集合"""
    black_set = set()
    black_file_path = os.path.join(FilePath, BLACK_LIST_FILENAME)
    if os.path.exists(black_file_path):
        with open(black_file_path, "r", encoding="utf-8") as f:
            for line in f:
                code = line.strip()
                if code:
                    black_set.add(code)
    return black_set


def append_to_blacklist(code):
    """追加单个代码到黑名单文件"""
    black_file_path = os.path.join(FilePath, BLACK_LIST_FILENAME)
    if not os.path.exists(FilePath):
        os.makedirs(FilePath, exist_ok=True)
    try:
        with open(black_file_path, "a", encoding="utf-8") as f:
            f.write(f"{code}\n")
    except Exception as e:
        print(f"❌写入黑名单失败 {code} , err:{str(e)}")


# ============================================================
# 三、安全保存（xlsx 优先，失败降级 csv）
# ============================================================
def safe_save_excel(data_list, save_path):
    if not data_list:
        print("⚠️ 无数据可保存，跳过")
        return False
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR, exist_ok=True)
    df = pd.DataFrame(data_list, columns=cols)
    saved = False
    # ----- xlsx -----
    try:
        tmp_xlsx = save_path + ".tmp.xlsx"
        with pd.ExcelWriter(tmp_xlsx, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="股票数据")
        if os.path.exists(save_path):
            os.remove(save_path)
        os.rename(tmp_xlsx, save_path)
        print(f"💾 已保存Excel，共 {len(df)} 条 -> {save_path}")
        saved = True
    except Exception as e:
        print(f"⚠️ 保存Excel失败（{e}），尝试保存为CSV...")
    # ----- csv 降级 -----
    if not saved:
        try:
            csv_path = save_path.rsplit(".", 1)[0] + ".csv"
            df.to_csv(csv_path, index=False, encoding="utf-8-sig")
            print(f"💾 已保存CSV（降级）共 {len(df)} 条 -> {csv_path}")
            saved = True
        except Exception as e2:
            print(f"❌ 保存CSV也失败！错误：{e2}")
    return saved


def load_done_from_save(save_path):
    """断点续传：从已保存的结果文件中读回已成功代码，避免重复请求"""
    done = set()
    for p in (save_path, save_path.rsplit(".", 1)[0] + ".csv"):
        if not os.path.exists(p):
            continue
        try:
            if p.endswith(".csv"):
                df = pd.read_csv(p, encoding="utf-8-sig")
            else:
                df = pd.read_excel(p)
            if "股票代码" in df.columns:
                done.update(df["股票代码"].astype(str).str.strip())
        except Exception as e:
            print(f"⚠️ 读取断点文件失败（{p}）：{e}")
    return done


# ============================================================
# 四、Futu 行情会话（统一封装，超时自动重建连接）
# ============================================================
class QuoteSession:
    def __init__(self):
        self.ctx = None
        self.reconnect()

    def reconnect(self):
        self.close()
        self.ctx = OpenQuoteContext(host=HOST, port=PORT)

    def close(self):
        if self.ctx is not None:
            try:
                self.ctx.close()
            except Exception:
                pass
            self.ctx = None

    def _call(self, codes):
        try:
            ret, data = self.ctx.get_market_snapshot(codes)
            time.sleep(SLEEP_INTERVAL)
            if ret == RET_OK:
                return "ok", data
            err = str(data)
            if any(k in err.lower() for k in BLACK_KEYWORDS):
                return "black", err
            return "error", err
        except Exception as e:
            err = str(e)
            if "Ret_TimeOut" in err:
                return "timeout", err
            return "error", err

    def get_snapshot(self, codes):
        timeout_cnt = 0
        for attempt in range(MAX_RETRY + 1):
            status, data = self._call(codes)
            if status == "timeout":
                timeout_cnt += 1
                print(f"!!! 通信超时（第{timeout_cnt}次），重建连接...")
                self.reconnect()
                time.sleep(RETRY_SLEEP)
                if timeout_cnt >= MAX_RETRY + 1:
                    return "error", data
                continue
            if status in ("ok", "black"):
                return status, data
            if attempt < MAX_RETRY:
                print(f"⏳ 请求失败，{RETRY_SLEEP}s 后重试 ({attempt + 1}/{MAX_RETRY})...")
                time.sleep(RETRY_SLEEP)
            else:
                return status, data
        return "error", "重试耗尽"


# ============================================================
# 五、处理状态与数据解析
# ============================================================
class State:
    def __init__(self, total):
        self.total = total
        self.rows = []
        self.done = set()
        self.processed = 0

    def add_black(self, code, reason):
        if code in self.done:
            return
        self.done.add(code)
        self.processed += 1
        append_to_blacklist(code)
        print(f"🚫 {code} -> {reason}，已加入黑名单")

    def add_row(self, code, last, low, high):
        if code in self.done:
            return
        self.done.add(code)
        self.processed += 1
        self.rows.append({
            "股票代码": code,
            "最新价格": last,
            "最新价格区间": "",
            "历史最低价格": low,
            "历史最低区间": "",
            "历史最高价格": high,
        })


def parse_snapshot_data(data, state, save_path):
    for idx, row in data.iterrows():
        code = str(row.get("code", "")).strip()
        if not code or code in state.done:
            continue
        last_price = row.get("last_price")
        if pd.isna(last_price) or last_price == 0:
            state.add_black(code, "价格无效")
            continue
        state.add_row(code, last_price,
                      row.get("lowest_history_price", None),
                      row.get("highest_history_price", None))
        remain = state.total - state.processed
        print(f"✅ {code} | 剩余:{remain} | 已收集:{len(state.rows)}")
        if len(state.rows) % AUTO_SAVE_INTERVAL == 0:
            safe_save_excel(state.rows, save_path)


def process_single(session, code, state, save_path):
    if code in state.done:
        return
    status, data = session.get_snapshot([code])
    if status == "ok":
        parse_snapshot_data(data, state, save_path)
    elif status == "black":
        state.add_black(code, "匹配黑名单关键词/无效代码")
    else:
        print(f"❌ {code} 多次请求失败，跳过")


# ============================================================
# 六、获取输入目录下所有文件
# ============================================================
def get_input_files(input_dir):
    """获取输入目录下的所有 CSV/XLSX 文件"""
    if not os.path.exists(input_dir):
        print(f"❌ 输入目录不存在: {input_dir}")
        return []
    
    valid_ext = {".csv", ".xlsx", ".xls"}
    files = []
    for f in os.listdir(input_dir):
        full_path = os.path.join(input_dir, f)
        if os.path.isfile(full_path) and os.path.splitext(f)[1].lower() in valid_ext:
            files.append(f)
    
    return sorted(files)


# ============================================================
# 七、处理单个文件
# ============================================================
def process_one_file(readfile):
    """处理单个输入文件"""
    pure_filename = os.path.basename(readfile)
    save_name = os.path.splitext(pure_filename)[0] + "demo.xlsx"
    save_path = os.path.join(OUTPUT_DIR, save_name)

    file_full_path = os.path.join(FilePath, readfile.lstrip("/"))

    print(f"\n{'='*60}")
    print(f"📄 正在处理文件: {pure_filename}")
    print(f"{'='*60}")

    # 核心：根据文件类型自动选择读取方式
    allstocklist = load_stock_codes_by_type(file_full_path)

    filename_upper = readfile.upper()
    market_name = "美股" if "US" in filename_upper else ("港股" if "HK" in filename_upper else "A股")
    print(f"📊 识别市场：{market_name}，共读取 {len(allstocklist)} 只股票")

    black_code_set = get_black_list()
    print(f"🔍 读取黑名单，共 {len(black_code_set)} 个无效代码")
    done_from_save = load_done_from_save(save_path)
    print(f"🔁 断点续传：已存在 {len(done_from_save)} 个已获取代码，将跳过")

    pending = [c for c in allstocklist
               if c not in black_code_set and c not in done_from_save]
    origin_count = len(allstocklist)
    skip_count = origin_count - len(pending)

    state = State(total=len(pending))
    estimate_min = (len(pending) / BATCH_SIZE * SLEEP_INTERVAL) / 60
    print(f"股票原始总数量：{origin_count}，跳过(黑名单+已获取)：{skip_count}，实际待请求：{len(pending)}")
    print(f"采用批量获取（每批{BATCH_SIZE}只），预估耗时 {estimate_min:.2f} 分钟\n")

    if not pending:
        print(f"✅ {pure_filename} 没有待处理的股票，跳过")
        return

    session = QuoteSession()
    start_time = time.time()

    for i in range(0, len(pending), BATCH_SIZE):
        batch = [c for c in pending[i:i + BATCH_SIZE] if c not in state.done]
        if not batch:
            continue
        print(f"\n📦 处理第 {i // BATCH_SIZE + 1} 批，共 {len(batch)} 只股票")

        status, data = session.get_snapshot(batch)
        if status == "ok":
            parse_snapshot_data(data, state, save_path)
        elif status == "black":
            print("↩️ 检测到无效代码，降级为逐个请求...")
            for code in batch:
                process_single(session, code, state, save_path)
        else:
            print(f"⚠️ 批量请求失败（{status}），降级为逐个请求...")
            for code in batch:
                process_single(session, code, state, save_path)

    session.close()

    safe_save_excel(state.rows, save_path)

    cost_sec = time.time() - start_time
    print(f"\n✅ {pure_filename} 处理完成！耗时：{cost_sec:.1f}s")
    print(f"   待处理：{state.total} | 成功：{len(state.rows)} | 黑名单/跳过：{state.total - len(state.rows)}")
    print(f"📁 数据文件：{save_path}")


# ============================================================
# 八、主流程
# ============================================================
def main():
    print(f"🎯 市场: {MARKET}")
    print(f"📂 输入目录: {FilePath}")
    print(f"📂 输出目录: {OUTPUT_DIR}")

    # 确保输出目录存在
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 获取所有输入文件
    input_files = get_input_files(FilePath)
    
    if not input_files:
        print(f"❌ 输入目录下没有 CSV/XLSX 文件: {FilePath}")
        return

    print(f"\n📋 检测到 {len(input_files)} 个待处理文件:")
    for i, f in enumerate(input_files, 1):
        print(f"  {i}. {f}")

    # 逐个处理
    success_count = 0
    fail_count = 0
    for readfile in input_files:
        try:
            process_one_file(readfile)
            success_count += 1
        except Exception as e:
            fail_count += 1
            print(f"\n❌ 文件 {readfile} 处理失败: {e}")
            continue

    print(f"\n{'='*60}")
    print(f"全部处理完成！")
    print(f"  成功: {success_count} 个文件")
    if fail_count > 0:
        print(f"  失败: {fail_count} 个文件")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
