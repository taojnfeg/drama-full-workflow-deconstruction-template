# -*- coding: utf-8 -*-
import os, sys, re, zipfile, docx
from docx.oxml.ns import qn

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "../.."))
sys.path.insert(0, os.path.join(BASE_DIR, "00_总控宪章与自动化引擎/scripts"))
from screenplay_gold_engine import verify_screenplay_8_assertions

EXPECTED_FILES = ['00_总控宪章与自动化引擎/00_《影视剧集从0到1全流程总控导航宪章（轻量路由入口）》.docx', '00_总控宪章与自动化引擎/01_按需加载全流程五阶段指导模块/模块01_立项初心（感动自己）、首集双引擎（高信息密度+新鲜度）与导演视听总纲指南.docx', '00_总控宪章与自动化引擎/01_按需加载全流程五阶段指导模块/模块02_开放世界生活场、底层运行铁律与立体人物关系四阶裂变（一人一档）指南.docx', '00_总控宪章与自动化引擎/01_按需加载全流程五阶段指导模块/模块03_高浓度场景空间拓扑（一景一档）与全局分集大纲（结构×情感填充物兼顾）指南.docx', '00_总控宪章与自动化引擎/01_按需加载全流程五阶段指导模块/模块04_双层细节工程（长伏笔+闲笔微镜像）、高信息量台词剧本与横向8列真导演分镜指南.docx', '00_总控宪章与自动化引擎/01_按需加载全流程五阶段指导模块/模块05_服化道顺场接戏统筹、三阶漏斗预告特辑与全流程100%完本交付核验指南.docx', '00_总控宪章与自动化引擎/02_《影视剧集从0到1全流程正向开发、视听转化与工业级交付总纲（完整合订本）》.docx', '00_总控宪章与自动化引擎/SKILL.md', '00_总控宪章与自动化引擎/scripts/audit_zero_errors.py', '00_总控宪章与自动化引擎/scripts/pub_layout_engine.py', '00_总控宪章与自动化引擎/scripts/screenplay_gold_engine.py', '01_全流程剧集开发标准母版库/00_项目总本与创作总论/00_《剧名》立项定位、首集双引擎（高信息密度+信息新鲜度）与故事核评估表（标准母版）.docx', '01_全流程剧集开发标准母版库/00_项目总本与创作总论/01_《剧名》全集专业文学剧本总本（标准母版）.docx', '01_全流程剧集开发标准母版库/00_项目总本与创作总论/02_《剧名》世界观设定、剧情大纲、戏剧结构与主创阐述总论（九编式标准母版）.docx', '01_全流程剧集开发标准母版库/01_专业文学剧本/EP01_《剧名》第一集《单集片名》专业文学剧本（高信息量台词与微动作交替标准母版）.docx', '01_全流程剧集开发标准母版库/02_导演分镜脚本与视听调度本/EP01_《剧名》第一集《单集片名》导演分镜脚本与视听调度本（横向宽幅8列标准母版）.docx', '01_全流程剧集开发标准母版库/03_剧作结构、信息密度节奏与视听阐述/EP01_《剧名》第一集《单集片名》剧作结构、信息密度节奏与视听阐述（标准母版）.docx', '01_全流程剧集开发标准母版库/04_按集全套完整交付包（一集一档：本集剧本+本集分镜+本集阐述）/EP01_第一集《单集片名》完整交付包/01_EP01_《剧名》第一集《单集片名》专业文学剧本.docx', '01_全流程剧集开发标准母版库/04_按集全套完整交付包（一集一档：本集剧本+本集分镜+本集阐述）/EP01_第一集《单集片名》完整交付包/02_EP01_《剧名》第一集《单集片名》导演分镜脚本与视听调度本（横向宽幅8列）.docx', '01_全流程剧集开发标准母版库/04_按集全套完整交付包（一集一档：本集剧本+本集分镜+本集阐述）/EP01_第一集《单集片名》完整交付包/03_EP01_《剧名》第一集《单集片名》剧作结构、信息密度节奏与视听阐述.docx', '01_全流程剧集开发标准母版库/05_核心人物一人一档/人物01_角色标准姓名（核心状态定位｜一人一档六章制·按集推进叙事标准母版）.docx', '01_全流程剧集开发标准母版库/06_核心场景一景一档/场景01_标准主场景名称（含核心子空间｜一景一档五章制·3列轴线调度标准母版）.docx', '01_全流程剧集开发标准母版库/07_服化道统筹与造型定妆顺场接戏总表/01_《剧名》全角色跨阶段造型定妆、伤效接戏与核心道具流转统筹总表（标准母版）.docx', '01_全流程剧集开发标准母版库/08_宣发视听转化与全流程演进复盘档案/01_《剧名》开播前三支预告片（先导·中段·终极）与四支幕后特辑分镜策划标准母版.docx', '01_全流程剧集开发标准母版库/08_宣发视听转化与全流程演进复盘档案/02_《剧名》从0到1全流程创作演进记录、跨集连续性核验台账与播后复盘档案（标准母版）.docx', 'README.md']

FORBIDDEN_PULL_JARGON = [
    "视听拉片本", "拉片核验", "02_过程库", "16宫格", "全人物正面面部核验",
    "必须经原片核实", "穿插原片真实", "原片中的呈现位置", "视听拉片参数表",
    "精确抄录原片", "对齐原片特写", "原片实证证件号码", "消除原片道具穿帮",
    "grab_frame_from_cos_url", "purify_process_jargon", "THE 项目示例 OF TIME"
]

FORBIDDEN_IP_HARDCODES = [
    "三人行", "赵左右", "荀未来", "赵了", "徐图图", "舒图", "王一漫", "罗爱文",
    "陈明", "戴娟", "赵东来", "刘美奕", "余止", "高尚", "长林", "贺羽端", "蔡雨潮",
    "白昼篇", "黄昏篇", "黑夜篇", "十元砂锅麻辣烫", "烤红薯铁炉", "三条鱼漫画屋",
    "卡奇滋", "八王寺", "人生是一场巨大的后知后觉", "《三个「我」》",
    "24集", "45分钟", "前8集", "后面20集", "EP08", "501", "502",
    "月牙", "左额角", "创可贴", "白敷贴", "酒红", "顶峰相见", "透明胶带",
    "99%", "帽兜", "读日历字", "手撕日历", "硬币", "失意成年人", "命定悖论", "跨时空",
    "有轨电车", "十余集", "3200K", "5600K", "6000K", "保温桶", "塔可夫斯基", "烟头", "大铁锤", "垃圾桶"
]

REQUIRED_METHODOLOGY_TERMS = [
    "感动自己", "情感填充物", "填充", "主心骨", "游离",
    "持续让观众关心", "信息密度", "猎奇", "新鲜度", "我没见过这样的东西"
]

def main():
    errors = []
    actual_files = []
    for r, dirs, files in os.walk(BASE_DIR):
        if ".git" in r or "__pycache__" in r:
            continue
        for f in files:
            actual_files.append(os.path.relpath(os.path.join(r, f), BASE_DIR))
    for ef in EXPECTED_FILES:
        if ef not in actual_files:
            errors.append(f"[缺失标准文件] {ef}")
    for af in actual_files:
        if af not in EXPECTED_FILES:
            errors.append(f"[多余未登记文件] {af}")

    all_corpus_text = ""
    for rel_p in sorted(actual_files):
        full_p = os.path.join(BASE_DIR, rel_p)
        if rel_p.endswith(".docx"):
            d = docx.Document(full_p)
            txt = "\n".join(p.text for p in d.paragraphs) + "\n" + "\n".join(c.text for tb in d.tables for r in tb.rows for c in r.cells)
            for ti, tb in enumerate(d.tables):
                for ri, row in enumerate(tb.rows):
                    if ri > 0:
                        trPr = row._tr.find(qn("w:trPr"))
                        if trPr is not None and trPr.findall(qn("w:cantSplit")):
                            errors.append(f"[跨页截断隐患] {rel_p} 表{ti} 行{ri} 包含 w:cantSplit")
            if "02_导演分镜脚本与视听调度本/" in rel_p or "02_EP01_" in rel_p:
                if len(d.tables) < 2 or len(d.tables[1].columns) != 8:
                    errors.append(f"[分镜表列数不符8列标准] {rel_p}")
            if "06_核心场景一景一档/" in rel_p:
                ch_cnt = sum(1 for p in d.paragraphs if re.match(r"^第[一二三四五]章", p.text.strip()))
                if ch_cnt != 5 or len(d.tables[1].columns) != 3:
                    errors.append(f"[场景档未达五章制或3列轴线表标准] {rel_p}: chapters={ch_cnt}, T2_cols={len(d.tables[1].columns)}")
            if "01_专业文学剧本/" in rel_p or "01_《剧名》全集专业文学剧本总本" in rel_p or "01_EP01_" in rel_p:
                verify_screenplay_8_assertions(full_p)
        else:
            txt = open(full_p, "r", encoding="utf-8").read()

        all_corpus_text += "\n" + txt
        if rel_p != "00_总控宪章与自动化引擎/scripts/audit_zero_errors.py":
            for bad_j in FORBIDDEN_PULL_JARGON:
                if bad_j in txt:
                    errors.append(f"[逆向拉片词残留] {rel_p} 包含 '{bad_j}'")
            for bad_ip in FORBIDDEN_IP_HARDCODES:
                if bad_ip in txt:
                    errors.append(f"[特定IP硬编码残留] {rel_p} 包含 '{bad_ip}'")

    for term in REQUIRED_METHODOLOGY_TERMS:
        if term not in all_corpus_text:
            errors.append(f"[核心方法论词条缺失] 全库未找到 '{term}'")

    if errors:
        for e in errors:
            print("ERROR:", e)
        sys.exit(1)
    print(f"[ALL 8 QUALITY GATES PASSED WITH ZERO ERRORS] Verified {len(actual_files)} pure-universal workflow files!")

if __name__ == "__main__":
    main()
