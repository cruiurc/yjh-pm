================
产品中心报告结构
================

2022-6-27

产品中心周报采取以下的结构规划

.. uml::
   
   @startmindmap
   + 产品中心报告
   ++ 中心概述
   +++ 进度概述
   ++++[#lightgreen] Table1：产品|PERT|警告|危险|进展
   +++++[#lightblue] PERT-mini
   +++ 人力成本概述
   ++++[#orange] chart1：项目|工时
   ++ 项目详情
   +++ 项目A
   ++++ 项目成员
   ++++[#lightblue] PERT
   ++++ 进展
   ++++ 备忘录
   ++++ 人力资源
   +++++[#orange] chart2：月|计划投入、实际投入
   ++++ 材料费用
   +++++[#orange] chart3：月|计划投入、实际投入
   ++++ 风险管理
   +++++[#lightgreen] Table2：WBS|工作包|计划开始时间|计划完成时间|风险|策略|进展
   ++++ 下一步计划
   +++++[#lightgreen] Table3：WBS|工作包|计划开始时间|计划完成时间
   +++ 项目B
   ++++ ...
   +++ ...
   @endmindmap

   

   
