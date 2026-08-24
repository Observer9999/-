import os
import json
import time
import requests
import datetime
import ddddocr
from playwright.sync_api import sync_playwright
from python_calamine import CalamineWorkbook
import subprocess
import random

# ================= 全局配置与本地调试开关 =================
LOCAL_DEBUG = False  # True: 本地跑(不git push), False: 云端跑(自动git push)
HISTORY_FILE = "meter_history.json"
CONFIG_FILE = "sites_config.json"
HISTORY_RETENTION_DAYS = 14

# 初始化 OCR
ocr = ddddocr.DdddOcr(show_ad=False)


def close_modal_if_exists(page, close_selector):
    """尝试关闭可能存在的遮挡弹窗"""
    if not close_selector:
        return False
    try:
        modal_close = page.locator(close_selector)
        # 如果关闭按钮存在且可见，就点击它
        if modal_close.count() > 0 and modal_close.first.is_visible():
            print(f"  🚪 检测到遮挡弹窗，正在自动关闭...")
            modal_close.first.click()
            page.wait_for_timeout(800)
            return True
    except:
        pass
    return False

# ================= 1. 历史状态管理 (JSON) =================
def load_history():
    """读取历史状态文件（增加空文件和格式错误的容错处理）"""
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if not content:
                    # 如果文件是空的，返回空字典
                    return {}
                return json.loads(content)
        except json.JSONDecodeError:
            # 如果文件格式损坏，打印警告并返回空字典
            print(f"️ {HISTORY_FILE} 内容格式错误，已重置为空状态。")
            return {}
    return {}


def save_and_push_history(history_data):
    """保存历史状态（本地调试时不执行 git push）"""
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history_data, f, indent=2, ensure_ascii=False)
    print(f"   📝 历史状态已保存到本地 {HISTORY_FILE}")

    if not LOCAL_DEBUG:
        try:
            # 【新增】配置 Git 用户信息，否则 push 会失败
            subprocess.run(["git", "config", "user.name", "github-actions[bot]"], check=True)
            subprocess.run(["git", "config", "user.email", "github-actions[bot]@users.noreply.github.com"], check=True)

            subprocess.run(["git", "add", HISTORY_FILE], check=True)
            subprocess.run(["git", "commit", "-m", f"Auto-update: meter history {datetime.date.today()}"], check=True)
            subprocess.run(["git", "push"], check=True)
            print("   ✅ 历史状态已自动推送到 GitHub")
        except subprocess.CalledProcessError as e:
            print(f"   ⚠️ Git 推送失败: {e}")
    else:
        print("   💡 [本地调试模式] 已跳过 Git 推送。请手动检查 meter_history.json 的内容。")


# ================= 2. 独立解析器库 (Parser Library) =================
def parse_znchaobiao_c_col(file_path, project_name, history):
    """
    【智能抄表网】专属解析逻辑
    规则：第6行起读C列，按天聚合取每日第一条读数。
    【升级】同步录入 history，实现全项目统一状态管理与滑动窗口维护。
    """
    workbook = CalamineWorkbook.from_path(file_path)
    project_data = {}

    # 确保该项目在 history 中存在
    if project_name not in history:
        history[project_name] = {}

    # 老网站是 T+0 数据，以今天为基准计算清理阈值
    today_dt = datetime.date.today()
    cutoff_date_dt = today_dt - datetime.timedelta(days=HISTORY_RETENTION_DAYS)
    cutoff_date_str = cutoff_date_dt.strftime("%Y-%m-%d")

    for sheet_name in workbook.sheet_names:
        sheet = workbook.get_sheet_by_name(sheet_name)
        data = sheet.to_python()
        daily_final_readings = {}

        # 1. 从 Excel 提取每日读数
        for row_idx in range(5, len(data)):
            row = data[row_idx]
            if len(row) > 2:
                try:
                    time_val = row[1]
                    reading_val = float(row[2])
                    date_str = time_val.split(" ")[0] if isinstance(time_val, str) else time_val.strftime("%Y-%m-%d")

                    # 取每天的第一条数据
                    if date_str not in daily_final_readings:
                        daily_final_readings[date_str] = reading_val
                except:
                    pass

        if daily_final_readings:
            # 2. 获取或初始化该电表的历史读数字典
            meter_history = history[project_name].get(sheet_name, {})
            readings = meter_history.get("readings", {})

            # 3. 【核心】用 Excel 权威数据覆盖/合并历史数据
            readings.update(daily_final_readings)

            # 4. 【核心】滑动窗口清理：只保留最近 N 天，防止 JSON 无限膨胀
            keys_to_delete = [date for date in readings.keys() if date < cutoff_date_str]
            for key in keys_to_delete:
                del readings[key]

            # 5. 更新 history，供持久化保存
            history[project_name][sheet_name] = {
                "readings": readings
            }

            # 6. 【完美统一】直接返回维护好的历史序列给分析节点
            # 这样老网站和新网站的分析数据源就完全一致了！
            project_data[sheet_name] = readings

            print(
                f"    📄 电表 [{sheet_name}] 聚合 {len(daily_final_readings)} 天数据，历史库当前维护 {len(readings)} 天。")

    return project_data


def parse_ns1886_diff(file_path, project_name, history):
    """
    【新网站/莱佛士】专属解析逻辑 (时间序列滑动窗口模式)
    """
    print(f"   ⚙️ 正在处理新网站文件: {file_path} (时间序列模式)")
    project_data = {}
    
    # 【核心修复】统一使用北京时间，杜绝 UTC 服务器导致的日期错位
    beijing_today = (datetime.datetime.utcnow() + datetime.timedelta(hours=8)).date()

    # J 列是前一天(T+1)的读数
    actual_date_dt = beijing_today - datetime.timedelta(days=1)
    actual_date_str = actual_date_dt.strftime("%Y-%m-%d")

    # 滑动窗口清理阈值
    cutoff_date_dt = actual_date_dt - datetime.timedelta(days=HISTORY_RETENTION_DAYS)
    cutoff_date_str = cutoff_date_dt.strftime("%Y-%m-%d")

    if project_name not in history:
        history[project_name] = {}

    try:
        workbook = CalamineWorkbook.from_path(file_path)
        sheet = workbook.get_sheet_by_index(0) 
        data = sheet.to_python()
        
        print(f"    [DEBUG] Excel 总行数: {len(data)}")
        
        # 从第 3 行开始 (索引 2)
        for row_idx in range(2, len(data)):
            row = data[row_idx]
            if not row: continue
            
            # 【调试】打印前 5 行，帮你确认表头和数据起始位置是否正确
            if row_idx < 7:
                print(f"    [DEBUG] Row {row_idx}: {row[:10]}... (总列数:{len(row)})")

            # 安全获取 B 列 (索引 1)
            meter_name = str(row[1]).strip() if len(row) > 1 else ""
            
            if not meter_name: 
                continue
            
            # 【修复】遇到合计行改为 continue (跳过)，而不是 break (停止)
            if "合计" in meter_name:
                print(f"    ⏭️ 跳过合计行: {meter_name}")
                continue 
                
            if len(row) <= 9:
                print(f"    ⚠️ 电表 {meter_name} 列数不足 10 列，跳过")
                continue

            try:
                current_reading = float(row[9]) # J列 (索引 9)
            except (ValueError, TypeError):
                print(f"    ⚠️ 电表 {meter_name} J列(索引9)格式错误: {row[9]}，跳过")
                continue

            # --- 以下为滑动窗口逻辑 (保持不变) ---
            meter_history = history[project_name].get(meter_name, {})
            if "last_total_diff" in meter_history and "readings" not in meter_history:
                last_date = meter_history.get("last_date", actual_date_str)
                readings = {last_date: meter_history["last_total_diff"]}
            else:
                readings = meter_history.get("readings", {})
            
            readings[actual_date_str] = current_reading
            
            keys_to_delete = [date for date in readings.keys() if date < cutoff_date_str]
            for key in keys_to_delete:
                del readings[key]
                
            history[project_name][meter_name] = {"readings": readings}
            project_data[meter_name] = readings
            
    except Exception as e:
        print(f"   ❌ 解析新网站文件失败: {e}")

    return project_data


# 解析器路由表
PARSER_MAP = {
    "parse_znchaobiao_c_col": parse_znchaobiao_c_col,
    "parse_ns1886_diff": parse_ns1886_diff
}



# nodes.py

# nodes.py

def process_single_account(acc, config, browser, global_history, log):
    """
    【核心重构】处理单个账号的完整流程：登录 -> 筛选 -> 下载 -> 解析
    返回: (是否成功, 项目名, 项目数据, 更新后的局部历史)
    """
    site_id = acc["site_id"]
    site = config["sites"].get(site_id)
    project_name = acc["project_name"]

    if not site:
        log(f"❌ 找不到站点配置: {site_id}")
        return False, project_name, None, None

    selectors = site["selectors"]
    features = site["features"]
    parser_func = PARSER_MAP.get(site["parser"])

    # 局部 History 隔离
    local_history = {project_name: global_history.get(project_name, {})}
    
    context = None
    try:
        log(f"\n{'=' * 60}")
        log(f" 正在处理项目: {project_name} (站点: {site['name']})")
        log(f"{'=' * 60}")

        context = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1920, "height": 1080},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()
        page.route("**/*.{png,jpg,jpeg,gif,svg,webp,ico,woff,woff2,ttf,eot}", lambda route: route.abort())

        # ================= 阶段 1：登录 =================
        log(f"  🌍 正在打开登录页: {site['login_url']}")
        goto_success = False
        for retry in range(10):
            try:
                page.goto(site["login_url"], timeout=120000, wait_until="domcontentloaded")
                goto_success = True
                break
            except Exception:
                log(f"  ⚠️ 打开网页超时 (第{retry + 1}次)，正在尝试刷新重试...")
                try:
                    page.reload(timeout=100000, wait_until="domcontentloaded")
                    goto_success = True
                    break
                except:
                    time.sleep(5)

        if not goto_success:
            log(f"  ❌ 无法打开登录页。")
            return False, project_name, None, None

        login_success = False
        max_retries = 10
        for attempt in range(1, max_retries + 1):
            log(f"  --- 第 {attempt} 次尝试登录 ---")
            try:
                page.fill(selectors["username_input"], acc["username"])
                page.fill(selectors["password_input"], acc["password"])

                if features.get("has_captcha"):
                    captcha_img = page.locator(selectors["captcha_img"])
                    captcha_img.wait_for(state="visible", timeout=8000)
                    captcha_text = ocr.classification(captcha_img.screenshot())
                    log(f"  🔍 OCR 识别结果: 【 {captcha_text} 】")
                    page.fill(selectors["captcha_input"], captcha_text)

                page.wait_for_timeout(300)
                page.click(selectors["login_btn"])
                page.wait_for_timeout(2500)

                if "login" not in page.url.lower():
                    log(f"  🎉 第 {attempt} 次登录成功！")
                    login_success = True
                    break
                else:
                    log(f"  ❌ 登录失败，准备刷新重试...")
                    page.reload()
                    page.wait_for_load_state("domcontentloaded")
            except Exception as e:
                log(f"  ⚠️ 登录尝试异常: {e}")
                time.sleep(2)

        if not login_success:
            log(f"  🚨 {project_name} 登录最终失败。")
            try:
                page.screenshot(path=f"login_fail_{project_name}.png")
            except:
                pass
            return False, project_name, None, None

        # ================= 阶段 2：跳转与筛选 =================
        close_modal_if_exists(page, selectors.get("close_modal_btn"))

        if features.get("has_menu"):
            log(f"  🧭 正在点击多级菜单...")
            for key in ["menu_level_1", "menu_level_2", "menu_level_3"]:
                if selectors.get(key):
                    try:
                        page.locator(selectors[key]).click(force=True)
                        page.wait_for_timeout(1000)
                    except Exception as e:
                        log(f"  ⚠️ 点击菜单 {key} 失败: {e}")
        else:
            log(f"  🧭 正在直接跳转: {site['data_url']}...")
            page.goto(site["data_url"], timeout=60000, wait_until="domcontentloaded")
            page.wait_for_load_state("networkidle", timeout=15000)

        close_modal_if_exists(page, selectors.get("close_modal_btn"))

        if features.get("has_complex_filter"):
            log(f"   正在执行多步筛选...")
            filters = selectors.get("filters", [])
            for i, f in enumerate(filters):
                trigger_selector = f.get("trigger")
                action = f.get("action", "first")
                option_selector = f.get("option_selector", ".el-select-dropdown__item")
                
                if trigger_selector:
                    for attempt in range(3):
                        try:
                            page.keyboard.press("Escape")
                            page.wait_for_timeout(300)
                            page.locator(trigger_selector).click(force=True)
                            page.wait_for_timeout(800)

                            if option_selector == ".el-select-dropdown__item":
                                dropdown = page.locator(".el-select-dropdown").filter(visible=True).last
                                target_options = dropdown.locator(option_selector)
                            else:
                                target_options = page.locator(option_selector)

                            visible_options = target_options.filter(visible=True)
                            if action == "last":
                                visible_options.last.click()
                            else:
                                visible_options.first.click()
                            break
                        except Exception as e:
                            if attempt == 2:
                                log(f"  ❌ 第 {i + 1} 步筛选最终失败")

        # ================= 阶段 3：查询与下载 =================
        if selectors.get("query_btn"):
            page.keyboard.press("Escape")
            page.mouse.click(10, 10)
            page.wait_for_timeout(500)
            page.locator(selectors["query_btn"]).click(force=True)
            page.wait_for_load_state("networkidle", timeout=15000)

        log(f"  📥 正在准备下载 Excel...")
        download_btn = selectors.get("download_btn")
        confirm_btn = selectors.get("confirm_download_btn")

        if not download_btn:
            log("   ⚠️ 未配置下载按钮选择器")
            return False, project_name, None, None

        try:
            if confirm_btn:
                page.locator(download_btn).click()
                page.wait_for_timeout(1000)
                with page.expect_download(timeout=60000) as download_info:
                    page.locator(confirm_btn).click()
            else:
                with page.expect_download(timeout=60000) as download_info:
                    page.locator(download_btn).click()

            download = download_info.value
            safe_filename = f"{project_name}_{download.suggested_filename}"
            save_path = os.path.join(os.getcwd(), safe_filename)
            download.save_as(save_path)
            log(f"  ✅ 文件已下载: {safe_filename}")
        except Exception as e:
            log(f"  ❌ 下载过程发生异常: {e}")
            return False, project_name, None, None

        # ================= 阶段 4：解析 =================
        if parser_func:
            project_data = parser_func(save_path, project_name, local_history)
            project_data["_meta"] = {
                "expected_lag_days": site.get("expected_lag_days", 0),
                "site_name": site.get("name", "未知站点")
            }
            log(f"  🎉 项目 {project_name} 解析完成")
            return True, project_name, project_data, local_history[project_name]
        else:
            log(f"  ❌ 未找到解析器: {site['parser']}")
            return False, project_name, None, None

    except Exception as e:
        log(f"  ❌ 项目 {project_name} 发生未知异常: {e}")
        return False, project_name, None, None
    finally:
        if context:
            try:
                context.close()
            except:
                pass



# nodes.py

def process_site_group(site_id, accounts, config, global_history):
    """
    【并行化核心】处理单个网站组（线程函数）
    每个网站组启动独立的 Playwright 实例，内部串行处理账号
    """
    # 初始化该线程的日志收集器
    thread_logs = []
    
    def log(msg):
        thread_logs.append(msg)
    
    # 该线程的结果收集
    results = {
        "success": [],      # [(project_name, project_data, project_history), ...]
        "failed": [],       # [project_name, ...]
        "retry_queue": []   # [acc, ...]
    }
    
    log(f"🌐 [线程] 启动网站组: {site_id}，包含 {len(accounts)} 个项目")
    
    try:
        # 【关键】每个线程启动独立的 Playwright 实例
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=not LOCAL_DEBUG, 
                slow_mo=500 if LOCAL_DEBUG else 0, 
                timeout=120000
            )
            log(f"✅ [线程] 浏览器启动成功: {site_id}")
            
            # 串行处理该网站下的所有账号（防风控）
            for acc in accounts:
                success, p_name, p_data, p_history = process_single_account(
                    acc, config, browser, global_history, log
                )
                
                if success:
                    results["success"].append((p_name, p_data, p_history))
                else:
                    results["failed"].append(p_name)
                    results["retry_queue"].append(acc)
                
                # 防风控
                time.sleep(random.uniform(2, 5))
            
            browser.close()
            log(f"🔒 [线程] 网站组 {site_id} 处理完成，浏览器已关闭")
            
    except Exception as e:
        log(f"❌ [线程] 网站组 {site_id} 发生异常: {e}")
        # 如果整个线程崩溃，将所有账号标记为失败
        for acc in accounts:
            results["failed"].append(acc["project_name"])
    
    return thread_logs, results
    

# ================= 3. 抽象的主抓取节点 =================
# nodes.py -> scrape_data_node

def scrape_data_node(state: dict):
    print("🚀 开始泛用性批量抓取数据 (并行化 + 全局重试)...")

    if not os.path.exists(CONFIG_FILE):
        print(f"❌ 找不到配置文件 {CONFIG_FILE}")
        return {"all_meter_data": {}, "login_success": False, "scrape_stats": {"total": 0, "success": 0, "failed_list": []}}

    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        config = json.load(f)

    history = load_history()
    all_meter_data = {}
    
    # 统计变量
    total_projects = len(config["accounts"])
    success_projects = []
    failed_projects = []
    
    # 【关键】按网站分组
    from collections import defaultdict
    site_groups = defaultdict(list)
    for acc in config["accounts"]:
        site_groups[acc["site_id"]].append(acc)
    
    # 【修复】使用 print 输出主流程信息
    print(f"📊 检测到 {len(site_groups)} 个网站，共 {total_projects} 个项目")
    print(f"📊 网站分布: {', '.join([f'{k}({len(v)}个)' for k, v in site_groups.items()])}")

    # ================= 第一轮：并行执行所有网站组 =================
    print("\n" + "="*20 + " 第一轮：并行抓取 " + "="*20)
    
    all_logs = []  # 收集所有线程的日志
    retry_queue = []  # 全局重试队列
    
    # 使用线程池并行执行
    from concurrent.futures import ThreadPoolExecutor, as_completed
    
    with ThreadPoolExecutor(max_workers=len(site_groups)) as executor:
        # 提交所有任务
        future_to_site = {
            executor.submit(process_site_group, site_id, accounts, config, history): site_id
            for site_id, accounts in site_groups.items()
        }
        
        # 等待所有任务完成
        for future in as_completed(future_to_site):
            site_id = future_to_site[future]
            try:
                thread_logs, results = future.result()
                all_logs.extend(thread_logs)
                
                # 处理成功的项目
                for p_name, p_data, p_history in results["success"]:
                    success_projects.append(p_name)
                    all_meter_data[p_name] = p_data
                    history[p_name] = p_history
                
                # 处理失败的项目
                failed_projects.extend(results["failed"])
                retry_queue.extend(results["retry_queue"])
                
            except Exception as e:
                print(f"❌ 网站组 {site_id} 线程异常: {e}")
                # 将该网站组的所有账号标记为失败
                for acc in site_groups[site_id]:
                    failed_projects.append(acc["project_name"])
                    retry_queue.append(acc)

    # ================= 第二轮：全局重试 =================
    if retry_queue:
        print("\n" + "="*20 + f" 第二轮：重试 {len(retry_queue)} 个失败项目 " + "="*20)
        # 清空失败列表，因为重试可能会成功
        failed_projects = []
        
        # 重试时按网站分组（保持防风控）
        retry_groups = defaultdict(list)
        for acc in retry_queue:
            retry_groups[acc["site_id"]].append(acc)
        
        with ThreadPoolExecutor(max_workers=len(retry_groups)) as executor:
            future_to_site = {
                executor.submit(process_site_group, site_id, accounts, config, history): site_id
                for site_id, accounts in retry_groups.items()
            }
            
            for future in as_completed(future_to_site):
                site_id = future_to_site[future]
                try:
                    thread_logs, results = future.result()
                    all_logs.extend(thread_logs)
                    
                    # 处理重试成功的项目
                    for p_name, p_data, p_history in results["success"]:
                        success_projects.append(p_name)
                        all_meter_data[p_name] = p_data
                        history[p_name] = p_history
                        print(f"✅ 重试成功: {p_name}")
                    
                    # 处理重试失败的项目
                    for p_name in results["failed"]:
                        failed_projects.append(p_name)
                        print(f"❌ 重试依然失败: {p_name}")
                    
                except Exception as e:
                    print(f"❌ 重试线程异常: {e}")
                    for acc in retry_groups[site_id]:
                        failed_projects.append(acc["project_name"])

    # ================= 统一输出日志 =================
    print("\n" + "="*60)
    print("📋 完整执行日志:")
    print("="*60)
    for log_line in all_logs:
        print(log_line)
    print("="*60)

    # 保存并推送历史状态
    save_and_push_history(history)

    # 构建最终统计
    scrape_stats = {
        "total": total_projects,
        "success": len(success_projects),
        "failed_list": failed_projects
    }

    # 【修复】使用 print 输出最终统计
    print(f"\n{'=' * 60}")
    print(f"✅ 抓取阶段结束！成功 {len(success_projects)}/{total_projects} 个项目")
    if failed_projects:
        print(f"❌ 最终失败项目: {', '.join(failed_projects)}")
    print(f"{'=' * 60}")
    
    is_any_success = len(success_projects) > 0
    return {"all_meter_data": all_meter_data, "login_success": is_any_success, "scrape_stats": scrape_stats}


# ================= 4. 分析节点 (保持原有健壮逻辑) =================
def analyze_data_node(state: dict):
    print("\n🔍 开始生成基于时间维度的报告...")
    all_data = state.get("all_meter_data", {})

    # 【新增】获取统计数据
    stats = state.get("scrape_stats", {})
    total = stats.get("total", 0)
    success = stats.get("success", 0)
    failed_list = stats.get("failed_list", [])

    if not all_data:
        return {"analyzed_report": "❌ 未获取到任何电表数据，请检查抓取日志。"}

    report_lines = ["# ⚡ 每日用电监控报告", ""]
     
    # 【新增】拼装统计摘要
    summary_lines = [f"**📊 抓取统计**：成功查询 {success}/{total} 个项目"]
    if failed_list:
        summary_lines.append(f"**❌ 失败项目**：{', '.join(failed_list)}")
    
    report_lines.extend(summary_lines)
    report_lines.append("") # 添加空行分隔
    report_lines.append("---") # 分割线
    report_lines.append("")

    has_any_anomaly = False
    today_str = (datetime.datetime.utcnow() + datetime.timedelta(hours=8)).strftime("%Y-%m-%d")

    for project_name, meters in all_data.items():
        project_anomalies = []
        normal_meters = []

        # 【新增】获取该项目的预期延迟天数，默认为 0
        meta = meters.get("_meta", {})
        expected_lag = meta.get("expected_lag_days", 0)

        for meter_name, daily_readings in meters.items():
            # 【新增】跳过元数据字典，只处理真实的电表数据
            if meter_name == "_meta":
                continue

            is_anomaly = False
            reason = ""
            sorted_dates = sorted(daily_readings.keys())

            if len(sorted_dates) < 2:
                is_anomaly = True
                reason = f"数据不足（仅 {len(sorted_dates)} 天记录，无法计算增量）"
            else:
                latest_date = sorted_dates[-1]
                prev_date = sorted_dates[-2]

                today_dt = datetime.datetime.strptime(today_str, "%Y-%m-%d")
                latest_dt = datetime.datetime.strptime(latest_date, "%Y-%m-%d")
                days_diff = (today_dt - latest_dt).days

                # 【核心修复】动态判断：实际延迟天数 > 预期延迟天数，即为断更！
                if days_diff > expected_lag:
                    is_anomaly = True
                    reason = f"数据断更（最新数据为 {latest_date}，滞后 {days_diff} 天，预期为 {expected_lag} 天）"
                else:
                    current_val = daily_readings[latest_date]
                    prev_val = daily_readings[prev_date]
                    increment = current_val - prev_val

                    history_incs = []
                    for i in range(1, len(sorted_dates)):
                        inc = daily_readings[sorted_dates[i]] - daily_readings[sorted_dates[i - 1]]
                        if inc > 0:
                            history_incs.append(inc)
                    avg_inc = sum(history_incs) / len(history_incs) if history_incs else 0

                    if increment <= 0:
                        is_anomaly = True
                        reason = f"增量为 {increment} (可能断电/故障)"
                    elif increment < 6:
                        is_anomaly = True
                        reason = f"增量过低 ({increment:.2f} < 6)"
                    elif avg_inc > 0 and increment < (avg_inc * 0.3):
                        is_anomaly = True
                        reason = f"增量骤降 ({increment:.2f})，平均 {avg_inc:.2f}"

            if is_anomaly:
                project_anomalies.append(f"- ⚠️ **{meter_name}**: {reason}")
                has_any_anomaly = True
            else:
                normal_meters.append(meter_name)

        report_lines.append(f"## 🏢 {project_name}")
        if project_anomalies:
            report_lines.append("### 异常电表：")
            report_lines.extend(project_anomalies)
            report_lines.append(f"\n*(该项目其余 {len(normal_meters)} 个电表正常)*")
        else:
            report_lines.append(f"✅ 全部正常 (共 {len(normal_meters)} 个电表)")
        report_lines.append("")

    report_lines.append("---")
    if has_any_anomaly:
        report_lines.append(
            "💡 **处理建议**：请优先检查异常电表的通讯模块、接线及现场设备。对于数据断更的电表，请检查采集器是否离线。")
    else:
        report_lines.append("✅ **总结**：今日所有项目用电数据正常，无异常波动。")

    final_report = "\n".join(report_lines)
    print("✅ 报告生成完毕。")
    return {"analyzed_report": final_report}


# ================= 5. 推送与失败处理节点 =================
def push_to_wechat_node(state: dict):
    """PushPlus 多人推送"""
    report = state.get("analyzed_report", "")
    print("\n📤 正在推送到微信...")

    tokens_str = os.getenv("PUSHPLUS_TOKEN")
    #tokens_str = os.getenv("TEST_TOKEN")
    if not tokens_str:
        print("❌ 未找到 PUSHPLUS_TOKEN")
        return {"status": "failed"}

    tokens = [token.strip() for token in tokens_str.split(",") if token.strip()]
    success_count = 0

    for token in tokens:
        print(f"  -> 正在向 Token: {token[:5]}... 发送推送")
        try:
            resp = requests.post("http://www.pushplus.plus/send", json={
                "token": token,
                "title": "⚡ 每日用电监控",
                "content": report,
                "template": "markdown"
            }, timeout=10).json()

            if resp.get("code") == 200:
                print("     ✅ 发送成功")
                success_count += 1
            else:
                print(f"     ❌ 发送失败: {resp.get('msg')}")
            time.sleep(1)  # 防频率限制
        except Exception as e:
            print(f"     ❌ 请求异常: {e}")

    print(f"\n📤 推送完成！共成功发送 {success_count}/{len(tokens)} 人。")
    return {"status": "success" if success_count > 0 else "failed"}


def handle_failure_node(state: dict):
    """处理失败"""
    print("🚨 任务失败。")
    return {"status": "failed"}
