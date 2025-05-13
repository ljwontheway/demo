import requests
from bs4 import BeautifulSoup
import json
import time
from pathlib import Path
import logging
from urllib.parse import urljoin

class AIBotCrawler:
    def __init__(self):
        self.base_url = "https://ai-bot.cn/"
        self.headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }
        self.output_dir = Path("output")
        self.output_dir.mkdir(exist_ok=True)
        self.setup_logging()

    def setup_logging(self):
        """设置日志"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('crawler.log', encoding='utf-8'),
                logging.StreamHandler()
            ]
        )

    def get_page(self, url):
        """获取页面内容"""
        try:
            response = requests.get(url, headers=self.headers)
            response.raise_for_status()
            return response.text
        except Exception as e:
            logging.error(f"获取页面失败 {url}: {str(e)}")
            return None

    def parse_categories(self):
        """解析左侧导航栏分类"""
        html = self.get_page(self.base_url)
        if not html:
            return []

        soup = BeautifulSoup(html, 'html.parser')
        categories = []

        # 查找主导航菜单 （在 sidebar-nav 中的 sidebar-menu）
        sidebar_menu = soup.find('div', class_='sidebar-menu flex-fill')
        if sidebar_menu:
            # 查找所有主菜单项（在 sidebar-menu-inner 中）
            menu_inner = sidebar_menu.find('div', class_='sidebar-menu-inner')
            if menu_inner:
                # 遍历所有导航项
                nav_items = menu_inner.find_all('li', class_='sidebar-item')
                for item in nav_items:
                    # 获取主分类名称
                    nav_link = item.find_all('a', class_='smooth')
                    if nav_link:
                        for nav in nav_link:
                            category = {
                                'name': nav.text.strip(),
                                'url': nav.get('href')
                            }
                        
                            if category['name']:
                                categories.append(category)
                                logging.info(f"找到分类: {category['name']},跳转地址：{category['url']}")

        if not categories:
            logging.error("未找到任何分类，请检查选择器是否正确")
            
        return categories

    def parse_item_detail(self, url):
        """解析详细内容页面"""
        html = self.get_page(url)
        if not html:
            return None

        soup = BeautifulSoup(html, 'html.parser')
        cards = []

        # 查找所有工具卡片
        rows = soup.findAll('div', class_='row io-mx-n2')
        for row in rows:
            segments = row.findAll('div', class_='url-card io-px-2 col-6 col-2a col-sm-2a col-md-2a col-lg-3a col-xl-6a col-xxl-6a')
            for segment in segments:
                try:
                    item = segment.find('a')
                    content = item.find('div', class_='text-sm overflowClip_1')
                    description = item.find('p')
                    card = {
                        'name': content.text.strip(),
                        'description': description.text.strip(),
                        'url':item.get('href') if item else '',
                    }
                    cards.append(card)
                    
                except Exception as e:
                    logging.error(f"解析卡片失败: {str(e)}")
                    continue

        return cards

    def crawl(self):
        """开始爬取"""
        try:
            # 获取所有分类
            categories = self.parse_categories()
            if not categories:
                logging.error("未找到分类信息")
                return

            results = []
            
            details = self.parse_item_detail(self.base_url)
            
            category_data = {
                'categories': categories,
                'subcategories': details
            }
            
           
            
            
            
            logging.info(details)
        

            results.append(category_data)

            # 保存结果
            self.save_results(results)
            logging.info("爬取完成")
            
        except Exception as e:
            logging.error(f"爬取过程出错: {str(e)}")

    def save_results(self, results):
        """保存爬取结果"""
        try:
            # 保存为JSON文件
            with open(self.output_dir / 'results.json', 'w', encoding='utf-8') as f:
                json.dump(results, f, ensure_ascii=False, indent=2)

            # 生成HTML报告
            self.generate_html_report(results)
            
        except Exception as e:
            logging.error(f"保存结果失败: {str(e)}")

    def generate_html_report(self, results):
        """生成HTML报告"""
        html = """
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <title>AI工具导航</title>
            <style>
                body { font-family: Arial, sans-serif; margin: 20px; }
                .category { margin-bottom: 30px; }
                .subcategory { margin: 20px 0; }
                .item { margin: 10px 0; padding: 10px; border: 1px solid #ddd; }
                .tag { background: #eee; padding: 2px 5px; margin-right: 5px; }
            </style>
        </head>
        <body>
        """

        for category in results:
            html += f"<div class='category'>"
            html += f"<h2>{category['name']}</h2>"
            
            for subcategory in category['subcategories']:
                html += f"<div class='subcategory'>"
                html += f"<h3>{subcategory['name']}</h3>"
                
                if subcategory['items']:
                    for item in subcategory['items']:
                        html += f"<div class='item'>"
                        html += f"<h4>{item['title']}</h4>"
                        html += f"<p>{item['description']}</p>"
                        if item['url']:
                            html += f"<p><a href='{item['url']}' target='_blank'>访问链接</a></p>"
                        if item['tags']:
                            html += "<p>"
                            for tag in item['tags']:
                                html += f"<span class='tag'>{tag}</span>"
                            html += "</p>"
                        html += "</div>"
                
                html += "</div>"
            
            html += "</div>"

        html += "</body></html>"

        with open(self.output_dir / 'report.html', 'w', encoding='utf-8') as f:
            f.write(html)

if __name__ == "__main__":
    crawler = AIBotCrawler()
    crawler.crawl()