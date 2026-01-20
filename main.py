import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import seaborn as sns
import scipy.stats as stats
import math
import tkinter as tk
from tkinter import filedialog
import os

# 配置绘图风格
plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号

# 使用Seaborn样式，但保留中文字体设置
sns.set_context("notebook")
sns.set_style("whitegrid", {"font.sans-serif": ['SimHei', 'Arial']})

# A4 尺寸 (Inches)
A4_SIZE = (11.69, 8.27) # Landscape usually better for charts, but tables might prefer Portrait
A4_PORTRAIT = (8.27, 11.69)
A4_LANDSCAPE = (11.69, 8.27)

def get_user_file_path():
    """弹出文件选择框获取文件路径"""
    try:
        if os.environ.get('HEADLESS_MODE'):
             return None
             
        root = tk.Tk()
        root.withdraw() # 隐藏主窗口
        
        current_dir = os.getcwd()
        file_path = filedialog.askopenfilename(
            initialdir=current_dir,
            title="请选择成绩汇总Excel文件",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        
        if not file_path:
            print("未选择文件，将使用默认测试路径(如果存在)。")
            return None
        return file_path
    except Exception as e:
        print(f"GUI启动失败(可能在无头环境): {e}")
        return None

def load_and_clean_data(file_path):
    """读取并清洗数据"""
    print(f"正在读取数据: {file_path}")
    try:
        df = pd.read_excel(file_path, sheet_name='班级全学科成绩汇总', engine='openpyxl')
    except Exception as e:
        print(f"读取Excel失败: {e}, 尝试默认引擎")
        df = pd.read_excel(file_path, sheet_name='班级全学科成绩汇总')

    # 删除冗余列 - 用户要求保留
    # cols_to_drop = ["序号", "准考证号", "自定义考号", "班级", "学生属性"]
    # df.drop(columns=[c for c in cols_to_drop if c in df.columns], inplace=True)

    # 删除最后一行合计行 - 用户要求删除
    if not df.empty:
         df = df.iloc[:-1]

    df.columns = df.columns.str.strip()
    
    # 强制类型转换
    if '姓名' in df.columns:
        df['姓名'] = df['姓名'].astype(str).str.strip()
        # 过滤姓名为空的行
        df = df[df['姓名'].notna() & (df['姓名'] != '') & (df['姓名'] != 'nan') & (df['姓名'] != 'None')]

    # 识别需要转换为数字的列
    cols_to_convert = [c for c in df.columns if '得分' in c or '校次' in c or '班次' in c or '总分' in c]
    if cols_to_convert:
        df[cols_to_convert] = df[cols_to_convert].apply(pd.to_numeric, errors='coerce')
    
    return df

def move_columns(df, cols_to_move, insert_after):
    """移动列位置"""
    cols = df.columns.tolist()
    remaining_cols = [col for col in cols if col not in cols_to_move]
    try:
        insert_idx = remaining_cols.index(insert_after) + 1
    except ValueError:
        insert_idx = 0
    new_cols = remaining_cols[:insert_idx] + cols_to_move + remaining_cols[insert_idx:]
    # 过滤掉不存在的列
    final_cols = [c for c in new_cols if c in df.columns]
    return df[final_cols]

def calculate_statistics(series, name):
    """计算统计指标"""
    if series.empty:
        return [name, 0, 0, 0, 0, 0, 0, 0]
    
    n = len(series)
    avg = series.mean()
    med = series.median()
    std = series.std()
    max_val = series.max()
    min_val = series.min()
    skew = series.skew()
    
    return [name, n, f"{avg:.2f}", f"{med:.2f}", f"{std:.2f}", 
            f"{max_val:.2f}", f"{min_val:.2f}", f"{skew:.2f}"]

def draw_stats_table(ax, stats_data, title):
    """(辅助) 在指定Axes上绘制统计表"""
    ax.axis('tight')
    ax.axis('off')
    cols = ["项目", "人数", "平均", "中位", "标差", "最高", "最低", "偏度"]
    table = ax.table(cellText=stats_data, colLabels=cols, cellLoc='center', loc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1, 1.5)
    ax.set_title(title)

def plot_summary_page(df, subjects_map, pdf):
    """绘制总体分析页"""
    print("正在绘制总体概览页...")
    # 改为 4x2 布局以容纳进退步表格
    fig, axes = plt.subplots(4, 2, figsize=(16, 24)) # 增加高度
    plt.subplots_adjust(hspace=0.5, wspace=0.3) # 增加垂直间距
    
    clean_data = df['总分'].dropna() if '总分' in df.columns else pd.Series()
    
    # 1. 总分密度分布
    ax = axes[0, 0]
    if not clean_data.empty:
        ax.hist(clean_data, bins=30, density=True, alpha=0.7, color='skyblue', edgecolor='black', label='频率')
        mu, sigma = clean_data.mean(), clean_data.std()
        x = np.linspace(clean_data.min(), clean_data.max(), 100)
        p = stats.norm.pdf(x, mu, sigma)
        ax.plot(x, p, 'r-', linewidth=2, label='正态拟合')
        ax.legend(loc='upper right')
        ax.set_title('总分密度分布图')

        # 添加统计表
        total_stats = [calculate_statistics(clean_data, "总分")]
        cols_total = ["项目", "人数", "平均", "中位", "标差", "最高", "最低", "偏度"]
        # 在图表上方添加表格
        the_table = ax.table(cellText=total_stats, colLabels=cols_total, 
                             loc='top', bbox=[0.0, 1.15, 1.0, 0.15]) # 调整位置不遮挡标题
        the_table.auto_set_font_size(False)
        the_table.set_fontsize(9)
        
    else:
        ax.text(0.5, 0.5, '无数据')
    
    # 2. 校次与进退步散点
    ax = axes[0, 1]
    if '校次' in df.columns and '校次进退步' in df.columns:
        ax.scatter(df['校次'], df['校次进退步'], alpha=0.6, c='grey', edgecolor='black')
        ax.axhline(0, color='red', linestyle='--')
        ax.set_xlabel('校次')
        ax.set_ylabel('校次进退步')
        ax.set_title('校次与进退步散点图')
    else:
        ax.text(0.5, 0.5, '无数据')

    # 3. 总分箱线图 (去色 + 散点)
    ax = axes[1, 0]
    if not clean_data.empty:
        # patch_artist=False 去除填充颜色
        ax.boxplot(clean_data, patch_artist=False, showmeans=True, 
                   meanprops=dict(marker='o', markerfacecolor='red', markeredgecolor='black'))
        # 增加抖动散点
        jitter = np.random.normal(1, 0.04, size=len(clean_data))
        ax.scatter(jitter, clean_data, alpha=0.5, color='grey', s=15, label='原始数据')
        ax.set_title('总分箱线图')
        ax.set_xticks([])
    else:
        ax.text(0.5, 0.5, '无数据')

    # 4. 各科目箱线图 (去色 + 散点)
    ax = axes[1, 1]
    boxplot_data = []
    boxplot_labels = []
    for sub_name, col_name in subjects_map.items():
        if col_name in df.columns:
            s_data = df[col_name].dropna()
            if not s_data.empty:
                boxplot_data.append(s_data)
                boxplot_labels.append(sub_name)
    
    if boxplot_data:
        ax.boxplot(boxplot_data, tick_labels=boxplot_labels, patch_artist=False, showmeans=True)
        # 增加抖动散点
        for i, data in enumerate(boxplot_data):
            jitter = np.random.normal(i + 1, 0.04, size=len(data))
            ax.scatter(jitter, data, alpha=0.5, color='grey', s=10)
        ax.set_title('各科目成绩分布箱线图')
    else:
        ax.text(0.5, 0.5, '无数据')

    # 5. 校次分段饼图
    ax = axes[2, 0]
    if '校次' in df.columns:
        rank_data = df['校次'].dropna()
        rank_bins = [0, 300, 600, 1200, float('inf')]
        rank_labels = ['前300名', '301-600名', '601-1200名', '1200名后']
        rank_cats = pd.cut(rank_data, bins=rank_bins, labels=rank_labels)
        rank_counts = rank_cats.value_counts(sort=False)
        if rank_counts.sum() > 0:
            ax.pie(rank_counts, labels=rank_counts.index, autopct='%1.1f%%', startangle=140, 
                   colors=sns.color_palette('pastel'), wedgeprops={'edgecolor':'black'})
            ax.set_title('校次分段占比')
    
    # 6. 进退步占比
    ax = axes[2, 1]
    if '校次进退步' in df.columns:
        prog_data = df['校次进退步'].dropna()
        if not prog_data.empty:
            counts = [prog_data.gt(0).sum(), prog_data.le(0).sum()]
            ax.pie(counts, labels=['进步', '退步或持平'], autopct='%1.1f%%', 
                   colors=['lightcoral', 'lightgreen'], wedgeprops={'edgecolor':'black'}, startangle=140)
            ax.set_title('校次进步/退步占比')
    
    # 7. 进退步前十统计表 (拆分为进步和退步)
    # 7.1 进步前十
    ax = axes[3, 0]
    ax.axis('tight')
    ax.axis('off')
    if '校次进退步' in df.columns:
        # 获取前10名 (进步，数值最大)
        top10_prog = df.nlargest(10, '校次进退步')[['姓名', '校次进退步']]
        if not top10_prog.empty:
            table_data = top10_prog.values
            cols = ['姓名', '进步幅度']
            table = ax.table(cellText=table_data, colLabels=cols, cellLoc='center', loc='center')
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.scale(1, 1.5)
            # 字体设置
            for (row, col), cell in table.get_celld().items():
                if row == 0:
                    cell.set_text_props(fontname='SimHei', weight='bold', color='white')
                    cell.set_facecolor('#2196F3')
                else:
                    cell.set_text_props(fontname='SimSun')
            ax.set_title('总分进步前十名')
        else:
             ax.text(0.5, 0.5, '无进步数据')
    else:
         ax.text(0.5, 0.5, '无进退步数据')

    # 7.2 退步前十
    ax = axes[3, 1]
    ax.axis('tight')
    ax.axis('off')
    if '校次进退步' in df.columns:
        # 获取前10名 (退步，数值最小)
        top10_reg = df.nsmallest(10, '校次进退步')[['姓名', '校次进退步']]
        # 也许可以按绝对值排序？如果不，直接展示负值即可
        # 这里直接展示
        if not top10_reg.empty:
            #为了美观，可以按升序排列(退步最大的排最前) -> nsmallest 已经是升序(最小在前)
            table_data = top10_reg.values
            cols = ['姓名', '退步幅度']
            table = ax.table(cellText=table_data, colLabels=cols, cellLoc='center', loc='center')
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.scale(1, 1.5)
            # 字体设置
            for (row, col), cell in table.get_celld().items():
                if row == 0:
                    cell.set_text_props(fontname='SimHei', weight='bold', color='white')
                    cell.set_facecolor('#f44336') # 红色
                else:
                    cell.set_text_props(fontname='SimSun')
            ax.set_title('总分退步前十名')
        else:
             ax.text(0.5, 0.5, '无退步数据')
    else:
         ax.text(0.5, 0.5, '无进退步数据')

    fig.suptitle('班级成绩总体分析概览', fontsize=16)
    pdf.savefig(fig)
    plt.close()

def plot_subject_page(df, subject_name, col_map, pdf):
    """绘制单科分析页"""
    print(f"正在绘制 {subject_name} 分析页...")
    score_col = col_map.get('score')
    rank_col = col_map.get('rank')
    progress_col = col_map.get('progress')

    if not score_col or score_col not in df.columns:
        return # 该科目无数据

    # 改为 4x2 布局以容纳更多图表 (新增一行放拆分后的表格)
    fig, axes = plt.subplots(4, 2, figsize=(16, 24))
    plt.subplots_adjust(hspace=0.4, wspace=0.3)
    
    # 1. 该科目校次占比
    ax = axes[0, 0]
    if rank_col and rank_col in df.columns:
        rank_data = df[rank_col].dropna()
        rank_bins = [0, 300, 600, 1200, float('inf')]
        rank_labels = ['前300名', '301-600名', '601-1200名', '1200名后']
        rank_cats = pd.cut(rank_data, bins=rank_bins, labels=rank_labels)
        rank_counts = rank_cats.value_counts(sort=False)
        if rank_counts.sum() > 0:
            ax.pie(rank_counts, labels=rank_counts.index, autopct='%1.1f%%', startangle=140, colors=sns.color_palette('pastel'))
            ax.set_title(f'{subject_name} 校次分段占比')
    else:
        ax.text(0.5, 0.5, '无校次数据')

    # 2. 该科目进退步占比
    ax = axes[0, 1]
    if progress_col and progress_col in df.columns:
        prog_data = df[progress_col].dropna()
        if not prog_data.empty:
            counts = [prog_data.gt(0).sum(), prog_data.le(0).sum()]
            ax.pie(counts, labels=['进步', '退步或持平'], autopct='%1.1f%%', colors=['lightcoral', 'lightgreen'])
            ax.set_title(f'{subject_name} 进退步占比')
    else:
        ax.text(0.5, 0.5, '无进退步数据')

    # 3. 校次与进退步散点图
    ax = axes[1, 0]
    if rank_col and progress_col and rank_col in df.columns and progress_col in df.columns:
        ax.scatter(df[rank_col], df[progress_col], alpha=0.6, c='orange', edgecolor='black')
        ax.axhline(0, color='red', linestyle='--')
        ax.set_xlabel('校次')
        ax.set_ylabel('校次进退步')
        ax.set_title(f'{subject_name} 校次 vs 进退步')
    else:
        ax.text(0.5, 0.5, '数据不全')

    # 4. 柱形图/直方图 + 正态分布曲线
    ax = axes[1, 1]
    score_data = df[score_col].dropna()
    if not score_data.empty:
        # 直方图
        ax.hist(score_data, bins=15, density=True, alpha=0.6, color='skyblue', edgecolor='black')
        # 正态拟合
        mu, sigma = score_data.mean(), score_data.std()
        x = np.linspace(score_data.min(), score_data.max(), 100)
        p = stats.norm.pdf(x, mu, sigma)
        ax.plot(x, p, 'r-', linewidth=2, label='正态拟合')
        
        ax.legend()
        ax.set_title(f'{subject_name} 成绩分布')
    else:
        ax.text(0.5, 0.5, '无成绩数据')

    # 5. 单科箱线图 (新增)
    ax = axes[2, 0]
    if not score_data.empty:
        # ... existing plotting code ...
        ax.boxplot(score_data, patch_artist=False, showmeans=True, 
                   meanprops=dict(marker='o', markerfacecolor='red', markeredgecolor='black'))
        # 增加抖动散点
        jitter = np.random.normal(1, 0.04, size=len(score_data))
        ax.scatter(jitter, score_data, alpha=0.5, color='grey', s=15)
        ax.set_title(f'{subject_name} 成绩箱线图')
        ax.set_xticks([])
    else:
        ax.text(0.5, 0.5, '无成绩数据')
    
    # 统计表 (2, 1)
    if not score_data.empty:
        sub_stats = [calculate_statistics(score_data, subject_name)]
        draw_stats_table(axes[2, 1], sub_stats, f'{subject_name} 统计指标')
    else:
        axes[2, 1].axis('off')
        axes[2, 1].text(0.5, 0.5, '无数据')
        
    # 6. 进步前十 (3, 0)
    ax = axes[3, 0]
    ax.axis('tight')
    ax.axis('off')
    if progress_col and progress_col in df.columns:
        # 获取前10名 (进步)
        top10_prog = df.nlargest(10, progress_col)[['姓名', progress_col]]
        if not top10_prog.empty:
            table_data = top10_prog.values
            cols = ['姓名', '进步幅度']
            table = ax.table(cellText=table_data, colLabels=cols, cellLoc='center', loc='center')
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.scale(1, 1.5)
            # 表头样式
            for (row, col), cell in table.get_celld().items():
                if row == 0:
                    cell.set_text_props(fontname='SimHei', weight='bold', color='white')
                    cell.set_facecolor('#2196F3')
                else:
                    cell.set_text_props(fontname='SimSun')
            ax.set_title(f'{subject_name} 进步前十名')
        else:
             ax.text(0.5, 0.5, '无进步数据')
    else:
         ax.text(0.5, 0.5, '无进退步数据')

    # 7. 退步前十 (3, 1)
    ax = axes[3, 1]
    ax.axis('tight')
    ax.axis('off')
    if progress_col and progress_col in df.columns:
        # 获取前10名 (退步)
        top10_reg = df.nsmallest(10, progress_col)[['姓名', progress_col]]
        if not top10_reg.empty:
            table_data = top10_reg.values
            cols = ['姓名', '退步幅度']
            table = ax.table(cellText=table_data, colLabels=cols, cellLoc='center', loc='center')
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.scale(1, 1.5)
            # 表头样式
            for (row, col), cell in table.get_celld().items():
                if row == 0:
                    cell.set_text_props(fontname='SimHei', weight='bold', color='white')
                    cell.set_facecolor('#f44336')
                else:
                    cell.set_text_props(fontname='SimSun')
            ax.set_title(f'{subject_name} 退步前十名')
        else:
             ax.text(0.5, 0.5, '无退步数据')
    else:
         ax.text(0.5, 0.5, '无进退步数据')

    fig.suptitle(f'{subject_name} 学科详细分析', fontsize=16)
    pdf.savefig(fig)
    plt.close()

def create_table_page(df, columns, title, pdf):
    """生成A4打印表格页，自动调整大小以尝试一页显示"""
    print(f"正在生成表格: {title}")
    
    # 筛选存在的列
    valid_cols = [c for c in columns if c in df.columns]
    data = df[valid_cols].copy()
    
    # 不分页，尝试在一页中显示所有数据
    # 计算需要的字体大小和行高
    total_rows = len(data) + 1 # +1 for header
    base_font_size = 10
    
    # 假设页面有效高度比例
    # 如果行数很多，缩小字体和行高
    if total_rows > 30:
        font_size = max(4, base_font_size * (30 / total_rows)) # 最小4号字
        scale_y = max(1.0, 1.5 * (30 / total_rows))
    else:
        font_size = base_font_size
        scale_y = 1.5
        
    # 创建图表
    fig, ax = plt.subplots(figsize=A4_PORTRAIT)
    ax.axis('tight')
    ax.axis('off')
    
    # 绘制表格
    table = ax.table(cellText=data.values,
                     colLabels=data.columns,
                     cellLoc='center',
                     loc='center')
    
    # 样式调整
    table.auto_set_font_size(False)
    table.set_fontsize(font_size)
    table.scale(1, scale_y) # x, y scale
    
    # 设置表头颜色和字体
    for (row, col), cell in table.get_celld().items():
        if row == 0:
            cell.set_facecolor('#4CAF50')
            cell.set_text_props(weight='bold', color='white', fontname='SimHei')
        else:
            cell.set_text_props(fontname='SimSun')
    
    # 自动调整列宽以适应页面
    # table.auto_set_column_width(col=list(range(len(data.columns))))
            
    # 让标题稍微高一点，避免重叠
    # ax.set_title(f"{title}", y=1.02, fontsize=12, fontname='SimHei')
    
    # 确保布局紧凑但不越界
    plt.tight_layout(rect=[0, 0.02, 1, 0.98])
    
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()

def create_top_performers_page(df, subjects_map, pdf):
    """生成进退步前十排名页"""
    print("正在生成进退步排名页...")
    
    # 准备要展示的类别：总分(用校次进退步) + 各科目
    categories = []
    
    # 校次进退步
    if '校次进退步' in df.columns:
        categories.append({'name': '总体校次', 'col': '校次进退步'})
        
    for sub_name, col_map in subjects_map.items():
        if col_map['progress'] and col_map['progress'] in df.columns:
            categories.append({'name': sub_name, 'col': col_map['progress']})
    
    # 动态计算子图布局
    n_cats = len(categories)
    n_cols = 3
    n_rows = math.ceil(n_cats / n_cols)
    
    fig, axes = plt.subplots(n_rows, n_cols, figsize=(A4_LANDSCAPE[0], A4_LANDSCAPE[1]*n_rows/2)) # 调整高度
    if n_rows == 1 and n_cols == 1: axes = np.array([axes]) 
    axes = axes.flatten() if isinstance(axes, np.ndarray) else [axes]
    
    for i, cat in enumerate(categories):
        ax = axes[i]
        col_name = cat['col']
        # 排序并取前10 (进退步越大越好)
        top10 = df.nlargest(10, col_name)[['姓名', col_name]]
        
        ax.axis('tight')
        ax.axis('off')
        table_data = top10.values
        cols = ['姓名', '进步幅度']
        
        table = ax.table(cellText=table_data, colLabels=cols, cellLoc='center', loc='center')
        table.auto_set_font_size(False)
        table.set_fontsize(8)
        table.scale(1, 1.5)
        
        # 表头样式
        table[(0, 0)].set_facecolor('#2196F3')
        table[(0, 1)].set_facecolor('#2196F3')
        table[(0, 0)].set_text_props(color='white')
        table[(0, 1)].set_text_props(color='white')
        
        ax.set_title(f"{cat['name']} 进步前十", fontsize=10)

    # 隐藏多余子图
    for j in range(n_cats, len(axes)):
        axes[j].axis('off')
        
    fig.suptitle('各学科及总分进步Top 10榜单', fontsize=16)
    plt.tight_layout()
    pdf.savefig(fig)
    plt.close()

def main():
    # 增加地址询问功能 (需求)
    file_path = get_user_file_path()
    if not file_path:
        # 如果用户取消，使用默认路径测试
        FilePath = "D:\\dev_env\\AnalyStudents\\data\\DataTable(1) clean.xlsx"
    else:
        FilePath = file_path
    
    # 确定输出路径 (同目录下)
    OutputPath = os.path.join(os.path.dirname(FilePath), "AnalyzeReport.pdf")
    
    # 1. 载入数据
    df = load_and_clean_data(FilePath)
    
    # 构建科目映射配置
    subjects_config = {
        '语文': {'score': '语文得分', 'rank': '语文校次', 'progress': '语文校次进退步'},
        '数学': {'score': '数学得分', 'rank': '数学校次', 'progress': '数学校次进退步'},
        '英语': {'score': '英语得分', 'rank': '英语校次', 'progress': '英语校次进退步'},
        '物理': {'score': '物理得分', 'rank': '物理校次', 'progress': '物理校次进退步'},
        '化学': {'score': '化学得分', 'rank': '化学校次', 'progress': '化学校次进退步'},
        '生物': {'score': '生物得分', 'rank': '生物校次', 'progress': '生物校次进退步'},
        '政治': {'score': '政治得分', 'rank': '政治校次', 'progress': '政治校次进退步'},
        '历史': {'score': '历史得分', 'rank': '历史校次', 'progress': '历史校次进退步'},
        '地理': {'score': '地理得分', 'rank': '地理校次', 'progress': '地理校次进退步'}
    }
    
    # 简单映射用于Summary页
    subjects_simple_map = {k: v['score'] for k,v in subjects_config.items()}

    # 开始生成PDF
    print(f"正在生成分析报告至: {OutputPath}")
    with PdfPages(OutputPath) as pdf:
        
        # 需求6: 合并生成一份PDF (已经通过 PdfPages 上下文管理器实现)
        
        # 1. 总体概览页
        plot_summary_page(df, subjects_simple_map, pdf)
        
        # 2. 总分成绩表 (需求2)
        # 将总分移到姓名后
        df_total = move_columns(df.copy(), ['总分', '校次', '校次进退步'], '姓名')
        
        # 确定要展示的列：包含姓名、总分相关、各科得分
        # 移除额外的属性列： "序号", "准考证号", "自定义考号", "班级", "学生属性"
        
        display_cols = ['姓名', '总分', '校次', '校次进退步']
        # 添加各科得分
        for sub, conf in subjects_config.items():
            if conf['score'] in df.columns:
                display_cols.append(conf['score'])
        
        # 不再添加 extra_cols
        # for c in extra_cols: ...
                
        create_table_page(df_total, display_cols, "班级总分及各科成绩表", pdf)
        
        # 3. 各学科分析页 (需求3) + 4. 各科成绩表 (需求4)
        for sub_name, config in subjects_config.items():
            # 生成学科分析图表页
            plot_subject_page(df, sub_name, config, pdf)
            
            # 生成学科成绩表页
            # 需求: 每一科的成绩相关数据列移动到姓名之后
            cols_to_move = [config['score'], config['rank'], config['progress']]
            # 过滤掉None或不存在的列
            cols_to_move = [c for c in cols_to_move if c and c in df.columns]
            
            if cols_to_move:
                df_sub = move_columns(df.copy(), cols_to_move, '姓名')
                # 仅展示该学科相关列 + 姓名 + 总分 + 总分进退步 (不含原始属性)
                sub_display_cols = ['姓名'] + cols_to_move + ['总分']
                if '校次进退步' in df.columns:
                    sub_display_cols.append('校次进退步')
                        
                create_table_page(df_sub, sub_display_cols, f"{sub_name}单科成绩表", pdf)
        
        # 5. 进退步前十页 (这里只保留总分的进退步Top10，各科的已经移入单科分析页)
        # 或者保留作为一个总览页？
        # 用户需求： "将各科进退步统计移动到各科分析页当中来" -> 指的是展示位置
        # 用户需求5(旧): "各筛选出前十... 生成在单独PDF页" -> 需求3(新) "同时将各科进退步统计移动到各科分析页当中来"
        # 这可能意味着不再需要单独的"各科进退步汇总页"，而是分散到了各科的分析页中。
        # 但"总分进退步前十"可能还需要？
        # 用户没有明确说删除旧的需求5，但说"移动到"。
        # 我们可以保留一个"总分进退步Top10"在这里，或者干脆只保留总分。
        # 函数 create_top_performers_page 可以修改为只生成 总分 的。
        
        # create_top_performers_page(df, {}, pdf) # 传空字典则只有总分
        # 但为了稳妥，我们还是生成一下总的Top 10页，不过只包含总分？
        # 既然已经分散到各页，也许汇总页可以只保留总分。
        pass
        
    print("所有任务完成！")

if __name__ == "__main__":
    main()
