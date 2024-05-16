import pandas as pd

# 读取三个Excel文件
physical_machine = pd.read_excel('物理机.xls')
virtual_machine = pd.read_excel('虚拟机 (2).xls')
operationObject = pd.read_excel('数字化人事档案系统项目（试点）\检查对象资产信息表-1.xlsx')

# 物理机表格的查询条件
exact_column_physical = '业务IP'
fuzzy_column_physical = '租户'

# 虚拟机表格的查询条件
exact_column_virtual = 'IP'
fuzzy_column_virtual = '租户'

# 获取精准查询关键词列
keywords_ip = operationObject['IP（必填）'].tolist()

# 获取模糊查询关键词列,需要我们填写
keywords_tenant = '数字'

# 将 NaN 替换为空字符串
operationObject.fillna('', inplace=True)

# 转换指定列的数据类型为字符串
operationObject['类型'] = operationObject['类型'].astype(str)
operationObject['实例名称'] = operationObject['实例名称'].astype(str)
operationObject['电源状态'] = operationObject['电源状态'].astype(str)
operationObject['备注'] = operationObject['备注'].astype(str)

# 循环遍历第一个表格的关键词，并进行查询
for keyword in keywords_ip:
    # 在物理机表格中同时进行精确查询和模糊查询
    matched_rows_physical = physical_machine[
        (physical_machine[exact_column_physical] == keyword)
        & (physical_machine[fuzzy_column_physical].str.contains(
            keywords_tenant))]
    # 如果物理机有匹配到数据，则添加到结果列表中
    if not matched_rows_physical.empty:
        condition = operationObject['IP（必填）'] == keyword  # 根据条件列值进行筛选
        type_t = '物理机'
        instance_name = matched_rows_physical['实例名称'].tolist()[0]
        power_status = matched_rows_physical['电源状态'].tolist()[0]
        operationObject.loc[condition, '类型'] = type_t
        operationObject.loc[condition, '实例名称'] = instance_name
        operationObject.loc[condition, '电源状态'] = power_status

    # 在虚拟机表格中同时进行精确查询和模糊查询
    matched_rows_virtual = virtual_machine[
        (virtual_machine[exact_column_virtual] == keyword)
        &
        (virtual_machine[fuzzy_column_virtual].str.contains(keywords_tenant))]
    # 如果虚拟机有匹配到数据，则添加到结果列表中
    if not matched_rows_virtual.empty:
        condition = operationObject['IP（必填）'] == keyword  # 根据条件列值进行筛选
        type_t = '虚拟机'
        instance_name = matched_rows_virtual['实例名称'].tolist()[0]
        power_status = matched_rows_virtual['电源状态'].tolist()[0]
        operationObject.loc[condition, '类型'] = type_t
        operationObject.loc[condition, '实例名称'] = instance_name
        operationObject.loc[condition, '电源状态'] = power_status

    if matched_rows_physical.empty and matched_rows_virtual.empty:
        condition = operationObject['IP（必填）'] == keyword
        remark = '未找到'
        operationObject.loc[condition, '备注'] = remark

# 保存回Excel文件    -处理过的

operationObject.to_excel('数字化人事档案系统项目（试点）\检查对象资产信息表-1-处理过的.xlsx',
                         index=False)
