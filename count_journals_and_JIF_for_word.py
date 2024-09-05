import os
import pandas as pd

def standardize_journal_name(name):
    """
    标准化期刊名：去掉 '& ' 和 '-'，并转换为小写
    """
    if pd.isna(name):
        return ''
    return name.replace('& ', '').replace('-', '').lower()

def get_highest_quartile(quartiles):
    """
    获取最高的 JIF Quartile 值，Q1 > Q2 > Q3 > Q4
    """
    quartile_order = ['Q1', 'Q2', 'Q3', 'Q4']
    quartiles_set = set(quartiles.dropna())  # 移除NaN值并转换为集合
    for q in quartile_order:
        if q in quartiles_set:
            return q
    return None

def process_journal_data(scie_path, jif_path, output_path):
    # 读取 SCI-E收录.xlsx 数据
    scie_df = pd.read_excel(scie_path)

    # 读取 期刊影响因子.xlsx 数据
    jif_df = pd.read_excel(jif_path)

    # 标准化 Source Title 和 Journal name 列
    scie_df['标准化期刊名'] = scie_df['Source Title'].apply(standardize_journal_name)
    jif_df['标准化期刊名'] = jif_df['Journal name'].apply(standardize_journal_name)

    # 统计 Source Title 出现次数
    journal_counts = scie_df['标准化期刊名'].value_counts().reset_index()
    journal_counts.columns = ['标准化期刊名', '论文数']

    # 合并两个数据表，使用标准化的期刊名进行匹配
    merged_df = pd.merge(journal_counts, jif_df, on='标准化期刊名', how='left')

    # 获取 JIF Quartile 值最高的记录
    grouped = merged_df.groupby('标准化期刊名').agg(
        论文数=('论文数', 'first'),
        影响因子2023年=('2023 JIF', 'max'),
        分区=('JIF Quartile', lambda x: get_highest_quartile(x))
    ).reset_index()

    # 恢复原始期刊名并排序
    grouped = pd.merge(grouped, scie_df[['标准化期刊名', 'Source Title']], on='标准化期刊名', how='left').drop_duplicates('标准化期刊名')

    # 转换 "影响因子2023年" 列为数字类型，并处理非数字值
    grouped['影响因子2023年'] = pd.to_numeric(grouped['影响因子2023年'], errors='coerce')

    # 按照影响因子2023年降序排列
    grouped.sort_values(by='影响因子2023年', ascending=False, inplace=True)

    # 添加序号列
    grouped.insert(0, '序号', range(1, len(grouped) + 1))

    # 计算论文数的总和
    total_papers = grouped['论文数'].sum()

    # 添加合计行
    total_row = pd.DataFrame({
        '序号': ['合计'],
        'Source Title': ['/'],
        '论文数': [total_papers],
        '影响因子2023年': ['/'],
        '分区': ['/']
    })

    grouped = pd.concat([grouped, total_row], ignore_index=True)

    # 使用 xlsxwriter 导出 Excel 文件，并设置字体为 Times New Roman
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        grouped[['序号', 'Source Title', '论文数', '影响因子2023年', '分区']].to_excel(writer, index=False, sheet_name='Sheet1')

        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']

        # 设置字体为 Times New Roman
        cell_format = workbook.add_format({'font_name': 'Times New Roman'})
        worksheet.set_column('A:B', None, cell_format)  # 设置 A 到 B 列的字体
        worksheet.set_column('C:E', None, cell_format)  # 设置 C 到 E 列的字体

        # 设置 C、D、E 列居中对齐
        center_format = workbook.add_format({'align': 'center', 'font_name': 'Times New Roman'})
        worksheet.set_column('C:E', None, center_format)

        # 设置 A 列左对齐
        left_align_format = workbook.add_format({'align': 'left', 'font_name': 'Times New Roman'})
        worksheet.set_column('A:A', None, left_align_format)

    print(f"结果已保存到 {output_path}")

def main():
    # 文件路径
    base_path = os.path.join('examples', '张健示例', 'SCI-E收录数据')
    scie_path = os.path.join(base_path, 'SCI-E收录.xlsx')
    jif_path = os.path.join(base_path, '期刊影响因子.xlsx')
    output_path = os.path.join('data_output', '2_SCI-E收录统计及影响因子与分区表_for_word.xlsx')

    # 处理数据
    process_journal_data(scie_path, jif_path, output_path)

if __name__ == "__main__":
    main()