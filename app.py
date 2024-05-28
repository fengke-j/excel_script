import pandas as pd

#每一个需要处理的excel表格
operation_filename = "D19海外勘探开发应用集成系统.xlsx"
# 获取模糊查询关键词列,需要我们填写
keywords_tenant = 'D9-海外勘探'

# 读取三个Excel文件
physical_machine = pd.read_excel('物理机.xls')
virtual_machine = pd.read_excel('虚拟机.xls')
operationObject = pd.read_excel('资产平台数据\\'+operation_filename)

# 检查并添加缺失的列
for col in ['类型', '实例名称', '备注', '电源状态']:
    if col not in operationObject.columns:
        operationObject[col] = ''

# 物理机表格的查询条件
exact_column_physical = '业务IP'
fuzzy_column_physical = '租户'

# 虚拟机表格的查询条件
exact_column_virtual = 'IP'
fuzzy_column_virtual = '租户'

# 获取精准查询关键词列
keywords_ip = operationObject['IP（必填）'].tolist()

# 将 NaN 替换为空字符串
operationObject.fillna('', inplace=True)

# 转换指定列的数据类型为字符串
operationObject['类型'] = operationObject['类型'].astype(str)
operationObject['实例名称'] = operationObject['实例名称'].astype(str)
operationObject['电源状态'] = operationObject['电源状态'].astype(str)
operationObject['备注'] = operationObject['备注'].astype(str)

# 在物理机表格中进行模糊查询
fuzzy_matched_rows_physical = physical_machine[
    physical_machine[fuzzy_column_physical].str.contains(keywords_tenant)]
# 在虚拟机表格中进行模糊查询
fuzzy_matched_rows_virtual = virtual_machine[
    virtual_machine[fuzzy_column_virtual].str.contains(keywords_tenant)]

# 初始化一个集合来记录已匹配的IP
matched_ips = set()

# 循环遍历第一个表格的关键词，并进行查询
for keyword in keywords_ip:
    # 初始化变量
    type_t = ''
    instance_name = ''
    power_status = ''
    remark = '未找到'

    # 在模糊物理机表格中进行精确查询
    exact_matched_rows_physical = fuzzy_matched_rows_physical[
        fuzzy_matched_rows_physical[exact_column_physical] == keyword]
    
    if not exact_matched_rows_physical.empty:
        type_t = '物理机'
        instance_name = exact_matched_rows_physical['实例名称'].tolist()[0]
        power_status = exact_matched_rows_physical['电源状态'].tolist()[0]
        remark = ''
        matched_ips.add(keyword)

    # 在虚拟机表格中进行精确查询
    exact_matched_rows_virtual = fuzzy_matched_rows_virtual[
        fuzzy_matched_rows_virtual[exact_column_virtual] == keyword]
    
    if not exact_matched_rows_virtual.empty:
        type_t = '虚拟机'
        instance_name = exact_matched_rows_virtual['实例名称'].tolist()[0]
        power_status = exact_matched_rows_virtual['电源状态'].tolist()[0]
        remark = ''
        matched_ips.add(keyword)

    condition = operationObject['IP（必填）'] == keyword
    operationObject.loc[condition, '类型'] = type_t
    operationObject.loc[condition, '实例名称'] = instance_name
    operationObject.loc[condition, '电源状态'] = power_status
    operationObject.loc[condition, '备注'] = remark

# 查找未匹配的物理机和虚拟机
unmatched_physical = fuzzy_matched_rows_physical[
    ~fuzzy_matched_rows_physical[exact_column_physical].isin(matched_ips)]
unmatched_virtual = fuzzy_matched_rows_virtual[
    ~fuzzy_matched_rows_virtual[exact_column_virtual].isin(matched_ips)]

# 将未匹配的数据添加到操作表的尾部
for _, row in unmatched_physical.iterrows():
    new_row = {
        'IP（必填）': row[exact_column_physical],
        '类型': '物理机',
        '实例名称': row['实例名称'],
        '电源状态': row['电源状态'],
        '备注': '表中没有'
    }
    operationObject = operationObject._append(new_row, ignore_index=True)

for _, row in unmatched_virtual.iterrows():
    new_row = {
        'IP（必填）': row[exact_column_virtual],
        '类型': '虚拟机',
        '实例名称': row['实例名称'],
        '电源状态': row['电源状态'],
        '备注': '表中没有'
    }
    operationObject = operationObject._append(new_row, ignore_index=True)

# 保存回Excel文件
operationObject.to_excel('资产平台数据-处理过的\\'+operation_filename, index=False)
