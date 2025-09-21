# -*- coding: utf-8 -*-
"""
HR Offer 自动生成工具 - 用户友好版
"""

import pandas as pd
from docxtpl import DocxTemplate, RichText
import os
import sys
import time
import traceback


def resource_path(relative_path):
    """获取资源的绝对路径，兼容PyInstaller打包与源码运行"""
    if os.path.isabs(relative_path):
        return relative_path

    # 可能的路径候选
    candidates = [
        *[os.path.join(base, relative_path) for base in [
            getattr(sys, '_MEIPASS', ''),
            os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else '',
            os.getcwd(),
            os.path.dirname(os.path.abspath(__file__))
        ] if base]
    ]

    for path in candidates:
        if os.path.exists(path):
            return path

    return os.path.abspath(relative_path)


def clean_string(value, default=''):
    """清理字符串，移除pandas的索引和dtype信息"""
    if pd.isna(value):
        return default

    result = str(value)

    # 清理pandas元数据
    if result.startswith('0 ') and len(result) > 2:
        result = result[2:]

    if any(x in result for x in ['dtype:', 'Name:']):
        result = '\n'.join(line for line in result.split('\n')
                           if 'dtype:' not in line and 'Name:' not in line)

    return result.replace('\\n', '\n').strip() or default


def build_rich_text_context(row, fields_config):
    """构建模板上下文，对需要提醒的字段使用特殊格式"""
    context = {}
    for key, config in fields_config.items():
        field, default, is_reminder = config
        value = row.get(field, default)

        # 检查是否需要特殊格式提醒
        if is_reminder and str(value).strip() == '请填写':
            # 创建特殊格式文字（红色加粗）
            rt = RichText()
            rt.add('请填写', color='FF0000', bold=True)
            context[key] = rt
        else:
            # 普通文本
            cleaned_value = clean_string(value, default) if isinstance(default, str) else value
            context[key] = cleaned_value

    return context


def build_normal_context(row, fields_config):
    """构建普通模板上下文"""
    context = {}
    for key, config in fields_config.items():
        field, default = config
        value = row.get(field, default)
        context[key] = clean_string(value, default) if isinstance(default, str) else value
    return context


def main():
    print("=== HR Offer 自动生成工具 ===\n正在启动...")

    # 配置文件路径
    CONFIG = {
        'excel': "candidate_data.xlsx",
        'offer_template': "offer_template.docx",
        'approval_template': "interview_approval.docx",
        'output_dir': "生成的文档"
    }

    # 使用资源路径
    paths = {k: resource_path(v) for k, v in CONFIG.items() if k != 'output_dir'}
    paths['output_dir'] = CONFIG['output_dir']

    # 检查必要文件
    missing = [(os.path.basename(path), path) for path in paths.values()
               if not os.path.exists(path) and path != paths['output_dir']]

    if missing:
        print("错误：找不到必要文件")
        for base, full in missing:
            print(f"- {base} [尝试路径: {full}]")

        print("\n请将以下文件与可执行程序放在同一目录：")
        for file in [CONFIG['excel'], CONFIG['offer_template'], CONFIG['approval_template']]:
            print(f"- {file}")
        time.sleep(2)
        return

    # 创建输出目录
    os.makedirs(paths['output_dir'], exist_ok=True)

    try:
        # 读取数据
        print("正在读取Excel数据...")
        offer_df = pd.read_excel(paths['excel'], sheet_name='offer信息')
        approval_df = pd.read_excel(paths['excel'], sheet_name='审批信息', dtype=str)

        print(f"找到 {len(offer_df)} 条候选人数据")
        
        # 显示所有候选人姓名供用户选择
        candidate_names = [clean_string(row.get('姓名', '')) for _, row in offer_df.iterrows() if clean_string(row.get('姓名', ''))]
        print(f"候选人列表: {', '.join(candidate_names)}")
        
        # 询问用户是否要处理所有候选人
        print(f"\n请选择处理模式:")
        print(f"1. 处理所有候选人 ({len(candidate_names)} 位)")
        print(f"2. 选择特定候选人")
        
        try:
            choice = input("请输入选择 (1 或 2，直接回车默认为1): ").strip()
            if choice == '2':
                print(f"\n请输入要处理的候选人姓名 (用逗号分隔，例如: 张三,李四):")
                selected_names_input = input("候选人姓名: ").strip()
                if selected_names_input:
                    selected_names = [name.strip() for name in selected_names_input.split(',') if name.strip()]
                    # 过滤数据框，只保留选中的候选人
                    offer_df = offer_df[offer_df['姓名'].isin(selected_names)]
                    print(f"已选择 {len(offer_df)} 位候选人进行处理")
                else:
                    print("未输入候选人姓名，将处理所有候选人")
            else:
                print("将处理所有候选人")
        except KeyboardInterrupt:
            print("\n用户取消操作")
            return
        except Exception as e:
            print(f"输入处理出错，将处理所有候选人: {str(e)}")

        # 字段映射配置 - Offer使用普通上下文
        OFFER_FIELDS = {
            'candidate_name': ('姓名', '未知候选人'),
            'occupation_name': ('任职职位', '未知职位'),
            'department_name': ('所属部门', '未知部门'),
            'address_name': ('办公地址', '未知地址'),
            'basic_salary': ('基本工资', 0),
            'bonus_salary': ('岗位工资', 0),
            'performance_salary': ('绩效工资', 0),
            'month_1': ('月', 0),
            'day_1': ('日', 0),
            'week_1': ('星期', '未知'),
            'probation_1': ('试用期比例', 0),
            'contact_1': ('HR', '未知'),
            'mobile_1': ('HR联系电话', 0),
            'offer_month': ('offer月', 0),
            'offer_day': ('offer日', 0),
        }

        # 字段映射配置 - 审批表使用特殊格式提醒上下文
        # 格式: (字段名, 默认值, 是否启用特殊格式提醒)
        APPROVAL_FIELDS = {
            'candidate_name': ('姓名', '未知候选人', False),
            'occupation_name': ('任职职位', '未知职位', False),
            'department_name': ('所属部门', '未知部门', False),
            'first_interview': ('初始面评', '', False),
            'interviewer_1': ('初始面试官', '', False),
            'second_interview': ('复试面评', '', False),
            'interviewer_2': ('复试面试官', '', False),
            'third_interview': ('终试面评', '', False),
            'interviewer_3': ('终面面试官', '', False),
            'leader_1': ('汇报对象', '请填写', True),
            'second_department': ('所属二级部门', '请填写', True),
            'total_salary': ('基本底薪', '请填写', True),
            'basic_salary1': ('基本薪资', '请填写', True),
            'bonus_salary1': ('岗位补助', '请填写', True),
            'performance_salary1': ('绩效工资', '请填写', True),
            'level_1': ('建议职级', '请填写', True),
            'department_leader': ('部门负责人', '请填写', True),
            'contact_2': ('招聘负责人', '请填写', True),
            'candidate_num': ('候选人联系电话', '请填写', True),
            'candidate_id': ('身份证号', '请填写', True),
            'channel_cate': ('招聘渠道', '请填写', True),
            'channel_detail': ('渠道备注', '请填写', True),
            'pre_salary': ('过往薪资描述', '', False),
            'expected_salary': ('期望薪资', '', False),
            'probation_months': ('试用期', '请填写', True),
            'probation_1': ('试用期比例', '请填写', True),
            'company_signed': ('签约主体', '请填写', True),
            'city_base': ('base地', '请填写', True),
        }

        success_counts = {'offer': 0, 'approval': 0}
        error_counts = {'offer': 0, 'approval': 0}
        total_candidates = len(offer_df)
        processed_count = 0

        print(f"开始批量处理 {total_candidates} 位候选人...\n")

        for index, offer_row in offer_df.iterrows():
            processed_count += 1
            candidate_name = clean_string(offer_row.get('姓名', ''))
            
            if not candidate_name:
                print(f"[{processed_count}/{total_candidates}] 跳过空姓名行")
                continue

            print(f"[{processed_count}/{total_candidates}] 正在处理: {candidate_name}")

            # 获取审批信息
            approval_data = approval_df[approval_df['姓名'] == candidate_name]
            if approval_data.empty:
                print(f"  ⚠️  跳过 {candidate_name}：找不到审批信息")
                continue

            # 为每个候选人创建新的模板实例，避免数据污染
            try:
                offer_tpl = DocxTemplate(paths['offer_template'])
                approval_tpl = DocxTemplate(paths['approval_template'])
            except Exception as e:
                print(f"  ❌ 加载模板失败: {str(e)}")
                continue

            try:
                # 构建上下文
                offer_context = build_normal_context(offer_row, OFFER_FIELDS)
                approval_context = build_rich_text_context(approval_data.iloc[0], APPROVAL_FIELDS)

                # 生成文件
                files_to_generate = [
                    (offer_tpl, f"全房通-员工录用通知书-{candidate_name}.docx", offer_context, 'offer'),
                    (approval_tpl, f"面试评估表+录用审批表-{candidate_name}.docx", approval_context, 'approval')
                ]

                candidate_success = True
                for template, filename, context, file_type in files_to_generate:
                    try:
                        output_path = os.path.join(paths['output_dir'], filename)
                        
                        # 检查文件是否已存在，如果存在则添加时间戳
                        if os.path.exists(output_path):
                            base_name, ext = os.path.splitext(filename)
                            timestamp = time.strftime("%Y%m%d_%H%M%S")
                            filename = f"{base_name}_{timestamp}{ext}"
                            output_path = os.path.join(paths['output_dir'], filename)
                        
                        template.render(context)
                        template.save(output_path)
                        print(f"  ✓ 已生成: {candidate_name}的{'Offer' if file_type == 'offer' else '审批表'}")
                        success_counts[file_type] += 1
                        
                    except Exception as e:
                        print(f"  ❌ 生成 {candidate_name} 的{'Offer' if file_type == 'offer' else '审批表'}失败: {str(e)}")
                        error_counts[file_type] += 1
                        candidate_success = False

                if candidate_success:
                    print(f"  🎉 {candidate_name} 处理完成")
                else:
                    print(f"  ⚠️  {candidate_name} 部分文件生成失败")

            except Exception as e:
                print(f"  ❌ 处理 {candidate_name} 时出错: {str(e)}")
                traceback.print_exc()
                error_counts['offer'] += 1
                error_counts['approval'] += 1

            print()  # 空行分隔

        # 输出结果
        print(f"\n🎉 批量处理完成！")
        print(f"📊 处理统计:")
        print(f"  - 总候选人数量: {total_candidates}")
        print(f"  - 成功生成Offer: {success_counts['offer']} 份")
        print(f"  - 成功生成审批表: {success_counts['approval']} 份")
        print(f"  - Offer生成失败: {error_counts['offer']} 份")
        print(f"  - 审批表生成失败: {error_counts['approval']} 份")
        print(f"📁 文件位置: '{paths['output_dir']}' 文件夹")
        
        if error_counts['offer'] > 0 or error_counts['approval'] > 0:
            print(f"\n⚠️  注意: 有部分文件生成失败，请检查错误信息并重试")
        else:
            print(f"\n✅ 所有文件生成成功！")

    except Exception as e:
        print(f"❌ 程序运行出错: {str(e)}")
        traceback.print_exc()

    time.sleep(2)


if __name__ == "__main__":
    main()