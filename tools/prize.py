import pandas as pd
import random
import tkinter as tk
from tkinter import messagebox, filedialog
import json
from pathlib import Path
import os

class LotterySystem:
    def __init__(self):
        self.employees = []  # 存储所有员工
        self.winners = {}    # 存储中奖人员
        self.prizes = {}     # 奖项配置
        self.load_config()   # 加载配置
        
    def load_config(self):
        """加载奖项配置"""
        config_path = Path('config.json')
        if config_path.exists():
            with open(config_path, 'r', encoding='utf-8') as f:
                self.prizes = json.load(f)
        else:
            # 默认配置
            self.prizes = {
                "特等奖": {"count": 1, "winners": []},
                "一等奖": {"count": 2, "winners": []},
                "二等奖": {"count": 3, "winners": []},
                "三等奖": {"count": 5, "winners": []}
            }
            self.save_config()

    def save_config(self):
        """保存奖项配置"""
        with open('config.json', 'w', encoding='utf-8') as f:
            json.dump(self.prizes, f, ensure_ascii=False, indent=2)

    def load_employees(self, file_path):
        """从Excel加载员工名单"""
        try:
            df = pd.read_excel(file_path)
            self.employees = df.to_dict('records')
            return len(self.employees)
        except Exception as e:
            print(f"加载员工名单失败: {str(e)}")
            return 0

    def draw(self, prize_level, count=1):
        """抽奖"""
        if prize_level not in self.prizes:
            return []
        
        # 获取未中奖的员工
        available_employees = [
            emp for emp in self.employees 
            if not any(emp in prize['winners'] for prize in self.prizes.values())
        ]
        
        # 如果可抽奖人数不足
        if len(available_employees) < count:
            return []
        
        # 随机抽取指定数量的员工
        winners = random.sample(available_employees, count)
        
        # 记录中奖人员
        self.prizes[prize_level]['winners'].extend(winners)
        
        return winners

    def reset(self):
        """重置抽奖结果"""
        for prize in self.prizes.values():
            prize['winners'] = []

    def export_results(self):
        """导出抽奖结果"""
        results = []
        for level, prize in self.prizes.items():
            for winner in prize['winners']:
                results.append({
                    '奖项': level,
                    '工号': winner['工号'],
                    '姓名': winner['姓名']
                })
        
        if results:
            df = pd.DataFrame(results)
            df.to_excel('抽奖结果.xlsx', index=False)
            return True
        return False

class PrizeUI:
    def __init__(self, root):
        self.root = root
        self.root.title("抽奖系统")
        
        # 设置窗口大小和位置
        window_width = 300
        window_height = 200
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 创建主框架
        main_frame = tk.Frame(self.root)
        main_frame.pack(expand=True)
        
        # 创建抽奖按钮
        self.draw_button = tk.Button(
            main_frame,
            text="开始抽奖",
            font=("微软雅黑", 14),
            command=self.draw_prize,
            width=15,
            height=2
        )
        self.draw_button.pack(pady=20)

    def draw_prize(self):
        # 抽奖逻辑
        winner = self.draw_random_winner()
        if winner:
            messagebox.showinfo("抽奖结果", f"恭喜 {winner} 中奖！")
        else:
            messagebox.showinfo("提示", "暂无可抽奖的人员")

    def draw_random_winner(self):
        # 实现随机抽奖逻辑
        # ... existing code ...
        pass

if __name__ == "__main__":
    app = PrizeUI(tk.Tk())
    app.root.mainloop()