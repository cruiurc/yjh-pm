import pandas as pd
import datetime as dt
import os
import matplotlib.pyplot as plt
import numpy as np

with open('./source/report/'+dt.date.today().strftime('%Y-%m-%d')+'.rst', 'w') as f:
    f.write('=======================================\n')
    f.write(dt.date.today().strftime('%Y-%m-%d')+'\n')
    f.write('=======================================\n')

for file in os.listdir('./data/'):
    if file.endswith('.xlsx'):
        product_name = os.path.splitext(file)[0]

        # 读取表数据
        df_member = pd.read_excel('./data/'+product_name+'.xlsx', sheet_name='成员').dropna()
        df_pert_issue = pd.read_excel('./data/'+product_name+'.xlsx', usecols=[0,1,2], sheet_name='PERT').dropna()
        df_pert_relation = pd.read_excel('./data/'+product_name+'.xlsx', usecols=[4,5], sheet_name='PERT').dropna()
        df_summary = pd.read_excel('./data/'+product_name+'.xlsx', sheet_name='进展').dropna()
        df_note = pd.read_excel('./data/'+product_name+'.xlsx', sheet_name='备忘录').dropna()
        df_plan = pd.read_excel('./data/'+product_name+'.xlsx', sheet_name='进度计划表').dropna()
        df_track = pd.read_excel('./data/'+product_name+'.xlsx', sheet_name='进度跟踪表').dropna()
        df_hr = pd.read_excel('./data/'+product_name+'.xlsx', sheet_name='人力资源表')
        df_cost = pd.read_excel('./data/'+product_name+'.xlsx', sheet_name='材料费用表')
        df_budget = pd.read_excel('./data/'+product_name+'.xlsx', sheet_name='产品成本表').dropna()


        # 输出.rst文本
        with open('./source/report/'+dt.date.today().strftime('%Y-%m-%d')+'.rst', 'a') as f:
            f.write(product_name+'\n')
            f.write('=======================================\n')
            f.write('成员\n')
            f.write('----------------\n\n')
            for i in df_member.index:
                title = df_member.loc[i, '职务']
                name = df_member.loc[i, '姓名']
                f.write(':'+title+':\n   '+name+'\n\n')
            f.write('\n')
            f.write('时间线\n')
            f.write('----------------\n\n')
            f.write('.. uml::\n\n')
            f.write('   @startuml\n')
            f.write('   left to right direction\n')
            f.write('   title PERT: '+product_name+'\n')
            
            for i in df_pert_issue.index:
                issue = df_pert_issue.loc[i, '里程碑']
                color = df_pert_issue.loc[i, '颜色']
                time =  df_pert_issue.loc[i, '计划完成时间'].strftime('%Y-%m-%d')
                f.write('   object '+issue+' #'+color+' {\n')
                f.write('   '+time+'\n')
                f.write('   }\n')
            for i in df_pert_relation.index:
                start = df_pert_relation.loc[i, '关系开始']
                end = df_pert_relation.loc[i, '关系结束']
                f.write('   '+start+' --> '+end+'\n')
            f.write('   @enduml\n\n')
            f.write('本周摘要\n')
            f.write('----------------\n\n')
            for i in df_summary.index:
                if df_summary.loc[i, '更新时间'] >= (dt.date.today() - dt.timedelta(days=7)):
                    title = df_summary.loc[i, '标题']
                    summary = df_summary.loc[i, '进展报告']
                    f.write(':'+title+':\n   '+summary+'\n')
            f.write('\n')
            f.write('决议、备忘录\n')
            f.write('----------------\n\n')
            for i in df_note.index:
                if df_note.loc[i, '更新时间'] >= (dt.date.today() - dt.timedelta(days=7)):
                    note = df_note.loc[i, '备忘录']
                    f.write('- '+note+'\n\n')
            f.write('\n')
            f.write('人力资源\n')
            f.write('----------------\n\n')
            f.write('.. figure:: '+'../_static/'+dt.date.today().strftime('%Y-%m-%d')+'-'+product_name+'-hr'+'.svg')
            f.write('\n\n')
            f.write('材料费用\n')
            f.write('----------------\n\n')
            f.write('.. figure:: '+'../_static/'+dt.date.today().strftime('%Y-%m-%d')+'-'+product_name+'-cost'+'.svg')            
            f.write('\n\n')
            f.write('产品成本\n')
            f.write('----------------\n\n')
            f.write('.. figure:: '+'../_static/'+dt.date.today().strftime('%Y-%m-%d')+'-'+product_name+'-budget'+'.svg')            
            f.write('\n\n')
            f.write('风险\n')
            f.write('----------------\n\n')
            f.write('.. list-table:: 风险跟踪表\n')
            f.write('   :header-rows: 1\n')
            f.write('   :widths: 4 15 15 15 25\n')
            f.write('   :stub-columns: 1\n\n')
            f.write('   *  -  风险级别\n      -  风险描述\n      -  风险影响\n      -  风险策略\n      -  关联工作包\n')
            for i in df_track.index:
                if df_track.loc[i, '更新时间'] >= (dt.date.today() - dt.timedelta(days=7)):
                    level = df_track.loc[i, '风险级别']
                    risk = df_track.loc[i, '风险描述']
                    effect = df_track.loc[i, '风险影响']
                    solution = df_track.loc[i, '风险策略']
                    wbs = str(df_track.loc[i, 'WBS'])
                    workpkg = df_track.loc[i, '工作包名称']
                    start = df_track.loc[i, '计划开始时间'].strftime('%Y-%m-%d')
                    end = df_track.loc[i, '计划结束时间'].strftime('%Y-%m-%d')
                    f.write('   *  -  '+level+'\n      -  '+risk+'\n      -  '+effect+'\n      -  '+solution+'\n      -  '+workpkg+'\n\n')
            f.write('下一步计划\n')
            f.write('----------------\n\n')
            f.write('.. list-table:: 下一步计划表\n')
            f.write('   :header-rows: 1\n')
            f.write('   :widths: 4 15 15 15\n')
            f.write('   :stub-columns: 1\n\n')
            f.write('   *  -  WBS\n      -  工作包名称\n      -  计划开始时间\n      -  计划结束时间\n')
            for i in df_plan.index:
                if df_plan.loc[i, '计划开始时间'] <= (dt.date.today() + dt.timedelta(days=7)):
                    wbs = str(df_plan.loc[i, 'WBS'])
                    workpkg = df_plan.loc[i, '工作包名称']
                    start = df_plan.loc[i, '计划开始时间'].strftime('%Y-%m-%d')
                    end = df_plan.loc[i, '计划结束时间'].strftime('%Y-%m-%d')
                    f.write('   *  -  '+wbs+'\n      -  '+workpkg+'\n      -  '+start+'\n      -  '+end+'\n\n')
            f.write('\n')
            f.close()
            # 输出chart
            # hr chart
            hr_mon = df_hr['月'].tolist()
            hr_plan = df_hr['计划投入'].tolist()
            hr_actual = df_hr['实际投入'].tolist()
            hr_plan_sum = df_hr['计划投入总计'].tolist()
            hr_actual_sum = df_hr['实际投入总计'].tolist()            

            fig = plt.figure()
            ax1=fig.add_subplot(212)
            ax1.plot(hr_mon, hr_plan_sum, label='cumulative plan')
            ax1.plot(hr_mon, hr_actual_sum, label='cumulative actual')
            ax1.legend()
            x = np.arange(len(hr_mon))
            ax2=fig.add_subplot(211)
            ax2.bar(x-.2, hr_plan, width=0.35, label='monthly plan')
            ax2.bar(x+.2, hr_actual, width=0.35, label='monthly actual')
            ax2.legend()
            plt.savefig('./source/_static/'+dt.date.today().strftime('%Y-%m-%d')+'-'+product_name+'-hr'+'.svg')

            # cost chart
            cost_mon = df_cost['月'].tolist()
            cost_plan = df_cost['计划费用'].tolist()
            cost_actual = df_cost['实际费用'].tolist()
            cost_plan_sum = df_cost['计划费用总计'].tolist()
            cost_actual_sum = df_cost['实际费用总计'].tolist()            

            fig = plt.figure()
            ax1=fig.add_subplot(212)
            ax1.plot(cost_mon, cost_plan_sum, label='cumulative plan')
            ax1.plot(cost_mon, cost_actual_sum, label='cumulative actual')
            ax1.legend()
            x = np.arange(len(cost_mon))
            ax2=fig.add_subplot(211)
            ax2.bar(x-.2, cost_plan, width=0.35, label='monthly plan')
            ax2.bar(x+.2, cost_actual, width=0.35, label='monthly actual')
            ax2.legend()
            plt.savefig('./source/_static/'+dt.date.today().strftime('%Y-%m-%d')+'-'+product_name+'-cost'+'.svg')

            # budget chart
            plan = df_budget['计划产品成本'].tolist()
            actual = df_budget['实际产品成本'].tolist()
            date = df_budget['更新时间'].tolist()

            fig = plt.figure()
            ax = fig.add_subplot(111)
            ax.plot(date, plan, label='plan', marker = 'o')
            ax.plot(date, actual, label='actual', marker = 'o')
            ax.legend()
            plt.savefig('./source/_static/'+dt.date.today().strftime('%Y-%m-%d')+'-'+product_name+'-budget'+'.svg')
            
