import os
from typing import TypedDict
from dotenv import load_dotenv
from langgraph.graph import StateGraph, END

# 导入 nodes.py 中的函数
from nodes import scrape_data_node, analyze_data_node, push_to_wechat_node, handle_failure_node

# 加载环境变量
load_dotenv()

# 1. 定义状态结构 (必须包含 all_meter_data)
class AgentState(TypedDict):
    all_meter_data: dict       # 存储抓取到的所有电表数据
    analyzed_report: str       # 存储大模型生成的报告
    login_success: bool        # 登录是否成功
    status: str                # 任务最终状态
    scrape_stats: dict         # 【新增】存储抓取统计信息 (total, success, failed_list)

# 2. 构建图
def build_graph():
    workflow = StateGraph(AgentState)

    # 添加节点
    workflow.add_node("scrape", scrape_data_node)
    workflow.add_node("analyze", analyze_data_node)
    workflow.add_node("push", push_to_wechat_node)
    workflow.add_node("fail", handle_failure_node)

    # 设置入口点
    workflow.set_entry_point("scrape")

    # 定义条件路由：根据 login_success 决定下一步
    def route_after_scrape(state: AgentState):
        if state.get("login_success"):
            return "analyze"
        else:
            return "fail"

    workflow.add_conditional_edges(
        "scrape",
        route_after_scrape,
        {"analyze": "analyze", "fail": "fail"}
    )

    # 分析后推送到微信
    workflow.add_edge("analyze", "push")
    
    # 推送后结束
    workflow.add_edge("push", END)
    workflow.add_edge("fail", END)

    return workflow.compile()

# 3. 执行图
if __name__ == "__main__":
    print("🤖 Agent 启动...")
    
    # 初始化状态
    initial_state = {
        "all_meter_data": {},
        "analyzed_report": "",
        "login_success": False,
        "status": "pending",
        "scrape_stats": {} 
    }
    
    app = build_graph()
    
    # 运行并获取最终状态
    final_state = app.invoke(initial_state)
    
    print(f"\n🏁 任务结束，最终状态: {final_state.get('status')}")
    if final_state.get("analyzed_report"):
        print("\n📢 最终播报内容:")
        print(final_state["analyzed_report"])
