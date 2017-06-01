# -*- coding:utf-8 -*-
import os
from mailmerge import MailMerge
from datetime import date

outfile = r"C:\Users\Administrator\Desktop\WorkRpt"
outname = "DayRpt-"+str(date.today())+".docx"
outpath = os.path.join(outfile,outname)
print outpath
template = r"D:\python\py27\python_docx\template.docx"
document = MailMerge(template)
print document.get_merge_fields()
today_work = u"1.准备空间分析展示PPT。2.继续完善yext的测试脚本，增加抑重转移的测试。3.自学一个ms word自动化处理的脚本"
finished_work = u"全部完成"
unfinished_work = u"无"
tomorrow_work = u"1.空间分析的宣讲。2.研究yext的接口GET LIST,TRACKING PIXEL, REVIEWS, CSS Selectors for Listings,写出需求文档。3.继续研究佳明写的strategy的结构和逻辑，写出代码结构图"
feels = u"今天对yext项目的代码又看了一遍，对input和output的内部的逻辑基本已经清楚，。对内部的一些错误机制还是不了解，有点是懂非懂。感觉java一层套一层，对层次理解不清楚，需要学习画层次图"
document.merge(
    date = format(date.today()),
    today_work = today_work,
    finished_work = finished_work,
    unfinished_work = unfinished_work,
    tomorrow_work = tomorrow_work,
    feels = feels
)
document.write(outpath)