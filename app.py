import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import io

# ================= 配置区域 (根据你的表格实际表头修改这里) =================
# 映射字典：把你的Excel表头映射为标准字段名
# 格式：'标准字段名': '你的Excel表头名'
COLUMN_MAP = {
    'ASIN': 'ASIN',
    'Title': '标题',  # 或者 '商品标题'
    'Price': '价格',
    'Monthly_Sales': '月销量',
    'Sales_Growth': '近30天销量增长率', # 确保是数值，比如 50% 或 0.5
    'Revenue_Change': '月销售额增长率', # 如果有这个字段
    'Price_Change': '价格变化',       # 如果有这个字段
    'Ratings': '评分数',
    'Launch_Date': '上架时间',
    'Brand': '品牌',
    'Weight': '重量',
    'Image_URL': '主图链接', # 或者 '图片'
    'SKU': 'SKU' # 卖家精灵导出通常有这个，没有则留空
}

# 需排除的品牌关键词 (可自行添加)
BLOCK_BRANDS = ['OXO', 'Ninja', 'KitchenAid', 'Keurig', 'AmazonBasics', 'Cuisinart']

# ========================================================================

def process_temu_selection(file_paths):
    """
    核心处理逻辑函数
    """
    all_data = []
    
    # 1. 读取三个文件并合并
    print("正在读取文件...")
    for file_path, source_name in file_paths:
        try:
            # 尝试读取Excel，跳过可能的标题行
            df = pd.read_excel(file_path)
            
            # 简单的列名清洗（去除空格）
            df.columns = [c.strip() for c in df.columns]
            
            # 标记来源
            df['Source_List'] = source_name
            all_data.append(df)
        except Exception as e:
            print(f"读取文件 {file_path} 失败: {e}")
            return None

    if not all_data:
        return None

    # 合并为一个大表
    full_df = pd.concat(all_data, ignore_index=True)
    
    # 2. 字段重命名 (标准化)
    # 反转映射以便重命名
    rename_dict = {v: k for k, v in COLUMN_MAP.items() if v in full_df.columns}
    full_df.rename(columns=rename_dict, inplace=True)
    
    # 确保关键数值列是数字类型
    numeric_cols = ['Price', 'Monthly_Sales', 'Ratings', 'Sales_Growth']
    for col in numeric_cols:
        if col in full_df.columns:
            # 转换百分比字符串 (如 "20%") 为数字
            if full_df[col].dtype == object:
                full_df[col] = full_df[col].astype(str).str.replace('%', '').str.replace(',', '')
            full_df[col] = pd.to_numeric(full_df[col], errors='coerce').fillna(0)

    # 3. 数据去重与清洗
    # 计算ASIN出现次数 (核心逻辑：重叠度)
    if 'ASIN' not in full_df.columns:
        print("错误：未找到 ASIN 列，请检查配置区域的 COLUMN_MAP")
        return None
        
    asin_counts = full_df['ASIN'].value_counts()
    
    # 保留唯一的ASIN用于分析，优先保留数据最全的一行
    unique_df = full_df.drop_duplicates(subset=['ASIN']).copy()
    
    # 4. 过滤逻辑 (Filter)
    print("正在清洗数据...")
    # 4.1 排除大牌
    if 'Brand' in unique_df.columns:
        pattern = '|'.join(BLOCK_BRANDS)
        unique_df = unique_df[~unique_df['Brand'].astype(str).str.contains(pattern, case=False, na=False)]
    
    # 4.2 排除价格过低 (假设小于$8难做)
    if 'Price' in unique_df.columns:
        unique_df = unique_df[unique_df['Price'] > 8]

    # 4.3 排除过重 (如果有重量列，且大于1000g)
    if 'Weight' in unique_df.columns:
         # 简单处理：假设有些单位是g，有些是kg，这里需根据实际数据调整，暂假设全是数字且为g
         unique_df = unique_df[unique_df['Weight'] < 1000]

    # 5. TPI 模型打分 (Scoring)
    print("正在计算 TPI 模型得分...")
    def calculate_score(row):
        score = 50 # 基础分
        
        asin = row['ASIN']
        
        # A. 重叠加分 (最重要的逻辑)
        count = asin_counts.get(asin, 1)
        if count == 2:
            score += 30
        elif count >= 3:
            score += 50
            
        # B. 销量加分
        if row.get('Monthly_Sales', 0) > 500:
            score += 10
            
        # C. 价格黄金区间加分 ($20 - $40)
        price = row.get('Price', 0)
        if 20 <= price <= 40:
            score += 10
        elif price < 15:
            score -= 10 # 价格太低扣分
            
        # D. 增长率加分
        if row.get('Sales_Growth', 0) > 50: # 假设数据是 50 代表 50%
            score += 10
            
        return score

    unique_df['TPI_Score'] = unique_df.apply(calculate_score, axis=1)
    
    # 6. 排序并取 Top 30
    top_30 = unique_df.sort_values(by='TPI_Score', ascending=False).head(30)
    
    # 7. 构建亚马逊链接
    top_30['Amazon_URL'] = 'https://www.amazon.com/dp/' + top_30['ASIN']
    
    # 8. 整理最终输出列
    output_cols = ['TPI_Score', 'ASIN', 'Amazon_URL', 'Title', 'Price', 
                   'Monthly_Sales', 'Sales_Growth', 'Ratings']
    
    # 添加用户特别要求的变化列 (如果存在)
    optional_cols = ['Revenue_Change', 'Price_Change', 'Image_URL', 'SKU']
    for col in optional_cols:
        if col in top_30.columns:
            output_cols.append(col)
            
    final_result = top_30[output_cols]
    
    return final_result

def visualize_results(df):
    """
    可视化 Top 30 的数据
    """
    if df is None or df.empty:
        return

    # 设置中文字体 (Colab中可能需要特殊处理，这里用通用设置)
    plt.rcParams['font.sans-serif'] = ['SimHei'] 
    plt.rcParams['axes.unicode_minus'] = False
    
    # 图表 1: Top 30 产品的销量增长率
    plt.figure(figsize=(12, 6))
    # 截取标题前15个字符以免太长
    short_titles = df['Title'].astype(str).str[:15] + '...'
    sns.barplot(x=df['Sales_Growth'], y=short_titles, palette='viridis')
    plt.title('Top 30 潜力爆款：近30天销量增长率 (%)')
    plt.xlabel('增长率 (%)')
    plt.ylabel('商品标题')
    plt.grid(axis='x', alpha=0.3)
    plt.tight_layout()
    plt.show()

    # 图表 2: 价格 vs 销量 气泡图 (气泡大小 = 得分)
    plt.figure(figsize=(10, 6))
    sns.scatterplot(data=df, x='Price', y='Monthly_Sales', 
                    size='TPI_Score', sizes=(50, 400), hue='TPI_Score', palette='coolwarm')
    plt.title('Top 30 商品分布：价格 vs 月销量 (气泡大小=推荐得分)')
    plt.xlabel('价格 ($)')
    plt.ylabel('月销量')
    plt.grid(True, alpha=0.3)
    plt.show()

# ================= 模拟运行 (实际使用时请替换下面的代码) =================
# 在 Google Colab 中，你需要运行以下代码来上传文件：
# from google.colab import files
# uploaded = files.upload()
# file_paths = []
# for fn in uploaded.keys():
#     if '增长' in fn: file_paths.append((fn, 'List_A_Growth'))
#     elif '评分' in fn: file_paths.append((fn, 'List_B_Rating'))
#     elif '时间' in fn: file_paths.append((fn, 'List_C_New'))
# df_result = process_temu_selection(file_paths)
# ======================================================================

# 如果你在本地运行，可以直接取消注释并填入你的文件名：
# file_paths = [
#     ('销量增长Top100.xlsx', 'List_A_Growth'),
#     ('评分数Top100.xlsx', 'List_B_Rating'),
#     ('上架时间Top100.xlsx', 'List_C_New')
# ]
# df_result = process_temu_selection(file_paths)
# if df_result is not None:
#     print("筛选完成！Top 5 商品预览：")
#     print(df_result.head())
#     visualize_results(df_result)
#     # 导出结果
#     df_result.to_excel("TEMU_Top30_Selection.xlsx", index=False)
#     print("结果已保存为 TEMU_Top30_Selection.xlsx")
