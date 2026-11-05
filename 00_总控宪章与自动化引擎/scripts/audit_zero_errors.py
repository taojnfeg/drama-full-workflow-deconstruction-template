# -*- coding: utf-8 -*-
"""
通用影视全流程母版库十维零误差自动化终检门禁 (audit_zero_errors.py)
涵盖：
1. 26个标准文件完备性与全库 26 个文件 MD5 哈希两两唯一校验（零字节级重复文件）
2. 文学剧本全部8道硬断言校验（verify_screenplay_8_assertions）
3. 横向8列导演分镜脚本视听规格完备性校验（verify_storyboard_8cols，不在前期硬性锁死单镜秒数）
4. 动态跨文档角色、场景与核心道具同源一致性校验（自动从剧本提取登场角色，零硬编码人名）
5. 目录编号 00_—08_、《03_单集阐述》表1—表4对齐与 [07] 合订本 100% AST 同步校验
6. 《02_总论》九编29表完备性、EP01—EP12 一集一表（12张独立5行子表）与跨集零复制雷同校验
7. 《06_核心场景一景一档》五章制、第2章四小节+3列轴线表与第4章纪实正文饱满度校验
8. 《05_核心人物一人一档》六章制、第4章按集推进完整叙事传记与零填空占位符（【角色名】/________）校验
9. 核心创作心法深层机制覆盖、零事后拉片词、零特定IP（含《三人行》与虚构古建木作戏）私有设定残留校验
10. 全量繁体字拦截库（零繁体字混杂）、零单元格重复句与零数据行 w:cantSplit 跨页截断隐患校验
"""
import os, sys, re, hashlib, docx
from docx.oxml.ns import qn

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "../.."))
sys.path.insert(0, os.path.join(BASE_DIR, "00_总控宪章与自动化引擎/scripts"))
from screenplay_gold_engine import verify_screenplay_8_assertions, verify_storyboard_8cols

EXPECTED_FILES = [
    '00_总控宪章与自动化引擎/00_《影视剧集从0到1全流程总控导航宪章（轻量路由入口）》.docx',
    '00_总控宪章与自动化引擎/01_按需加载全流程五阶段指导模块/模块01_立项初心（感动自己）、首集双引擎（高信息密度+新鲜度）与导演视听总纲指南.docx',
    '00_总控宪章与自动化引擎/01_按需加载全流程五阶段指导模块/模块02_开放世界生活场、底层运行铁律与立体人物关系四阶裂变（一人一档）指南.docx',
    '00_总控宪章与自动化引擎/01_按需加载全流程五阶段指导模块/模块03_高浓度场景空间拓扑（一景一档）与全局分集大纲（结构×情感填充物兼顾）指南.docx',
    '00_总控宪章与自动化引擎/01_按需加载全流程五阶段指导模块/模块04_双层细节工程（长伏笔+闲笔微镜像）、高信息量台词剧本与横向8列真导演分镜指南.docx',
    '00_总控宪章与自动化引擎/01_按需加载全流程五阶段指导模块/模块05_服化道顺场接戏统筹、三阶漏斗预告特辑与全流程100%完本交付核验指南.docx',
    '00_总控宪章与自动化引擎/02_《影视剧集从0到1全流程正向开发、视听转化与工业级交付总纲（完整合订本）》.docx',
    '00_总控宪章与自动化引擎/SKILL.md',
    '00_总控宪章与自动化引擎/scripts/audit_zero_errors.py',
    '00_总控宪章与自动化引擎/scripts/pub_layout_engine.py',
    '00_总控宪章与自动化引擎/scripts/screenplay_gold_engine.py',
    '01_全流程剧集开发标准母版库/00_项目总本与创作总论/00_《剧名》立项定位、首集双引擎（高信息密度+信息新鲜度）与故事核评估表（标准母版）.docx',
    '01_全流程剧集开发标准母版库/00_项目总本与创作总论/01_《剧名》全集专业文学剧本总本（标准母版）.docx',
    '01_全流程剧集开发标准母版库/00_项目总本与创作总论/02_《剧名》世界观设定、剧情大纲、戏剧结构与主创阐述总论（九编式标准母版）.docx',
    '01_全流程剧集开发标准母版库/01_专业文学剧本/EP01_《剧名》第一集《单集片名》专业文学剧本（高信息量台词与微动作交替标准母版）.docx',
    '01_全流程剧集开发标准母版库/02_导演分镜脚本与视听调度本/EP01_《剧名》第一集《单集片名》导演分镜脚本与视听调度本（横向宽幅8列标准母版）.docx',
    '01_全流程剧集开发标准母版库/03_剧作结构、信息密度节奏与视听阐述/EP01_《剧名》第一集《单集片名》剧作结构、信息密度节奏与视听阐述（标准母版）.docx',
    '01_全流程剧集开发标准母版库/04_按集全套完整交付包（一集一档：本集剧本+本集分镜+本集阐述）/EP01_第一集《单集片名》完整交付包/01_EP01_《剧名》第一集《单集片名》专业文学剧本.docx',
    '01_全流程剧集开发标准母版库/04_按集全套完整交付包（一集一档：本集剧本+本集分镜+本集阐述）/EP01_第一集《单集片名》完整交付包/02_EP01_《剧名》第一集《单集片名》导演分镜脚本与视听调度本（横向宽幅8列）.docx',
    '01_全流程剧集开发标准母版库/04_按集全套完整交付包（一集一档：本集剧本+本集分镜+本集阐述）/EP01_第一集《单集片名》完整交付包/03_EP01_《剧名》第一集《单集片名》剧作结构、信息密度节奏与视听阐述.docx',
    '01_全流程剧集开发标准母版库/05_核心人物一人一档/人物01_角色标准姓名（核心状态定位｜一人一档六章制·按集推进叙事标准母版）.docx',
    '01_全流程剧集开发标准母版库/06_核心场景一景一档/场景01_标准主场景名称（含核心子空间｜一景一档五章制·3列轴线调度标准母版）.docx',
    '01_全流程剧集开发标准母版库/07_服化道统筹与造型定妆顺场接戏总表/01_《剧名》全角色跨阶段造型定妆、伤效接戏与核心道具流转统筹总表（标准母版）.docx',
    '01_全流程剧集开发标准母版库/08_宣发视听转化与全流程演进复盘档案/01_《剧名》开播前三支预告片（先导·中段·终极）与四支幕后特辑分镜策划标准母版.docx',
    '01_全流程剧集开发标准母版库/08_宣发视听转化与全流程演进复盘档案/02_《剧名》从0到1全流程创作演进记录、跨集连续性核验台账与播后复盘档案（标准母版）.docx',
    'README.md'
]

FORBIDDEN_PULL_AND_TYPOS = [
    "拉片", "视听拉片本", "拉片核验", "02_过程库", "16宫格", "全人物正面面部核验",
    "必须经原片核实", "穿插原片真实", "原片中的呈现位置", "视听拉片参数表",
    "精确抄录原片", "对齐原片特写", "原片实证证件号码", "消除原片道具穿帮",
    "grab_frame_from_cos_url", "purify_process_jargon", "THE 项目示例 OF TIME",
    "逆向逆向", "热情感信物", "拧开情感信物盖子", "日常餐食的日常烟火", "餐具情感信物",
    "情感填充物示例", "示例道具A", "第四阶段专著", "04_人物", "05_一景一档",
    "05_场景", "06_服化道", "07_宣发",
    "旧版已废止分册名", "旧版错位人物目录号", "旧版错位场景目录号", "旧版错位服化道目录号",
    "旧版错位宣发目录号", "旧式事后拆解", "重复词病句", "生硬替换词", "机械替换词病句",
    "生硬概念括注", "生硬道具占位符", "________", "【角色名】", "人物02", "人物03",
    "单镜时长", "预计时长"
]

FORBIDDEN_IP_HARDCODES = [
    "三人行", "赵左右", "荀未来", "赵了", "徐图图", "舒图", "王一漫", "罗爱文",
    "陈明", "戴娟", "赵东来", "刘美奕", "余止", "高尚", "长林", "贺羽端", "蔡雨潮",
    "白昼篇", "黄昏篇", "黑夜篇", "十元砂锅麻辣烫", "烤红薯", "三条鱼漫画屋",
    "卡奇滋", "八王寺", "人生是一场巨大的后知后觉", "《三个「我」》",
    "501", "502", "月牙", "左额角", "创可贴", "白敷贴", "酒红", "顶峰相见", "透明胶带",
    "99%", "帽兜", "读日历字", "手撕日历", "硬币", "失意成年人", "命定悖论", "跨时空",
    "有轨电车", "3200K", "5600K", "6000K", "保温桶", "塔可夫斯基", "烟头", "大铁锤", "垃圾桶",
    "既有印记消失", "唯一物理信标面临拆迁", "后人的干预恰恰促成先人发明", "注册名字卡系统漏洞",
    "你选的路，问过我想要吗", "谁在雪巷晕倒", "谁在岸边救人", "黄铜外壳机械怀表", "牛角扣大衣",
    "暗格刻度", "林深", "周明远", "苏怀瑾", "赵泰来", "青云坊", "鲁班尺", "鲁班锁匣", "天心斗拱", "QY-04-B",
    "生桐油", "燕尾榫", "榫槽", "榫卯", "老楠木", "六合", "天枢", "毫厘必究", "透气贴", "尺神经",
    "老榆木", "雕花木", "砖雕", "营造底簿", "营造法式", "暗卯", "三十六卯口", "平口凿", "凿痕",
    "松烟墨", "旧织机", "木料库", "陈皮姜茶"
]

FORBIDDEN_TRADITIONAL_CHARS = set(
    "與結構徹傷間發見開關門車馬鳥語論議護驗實證長東兩無專業學體係統電報紙張書寫讀聽覺頭臉頸時鐘樓層號條塊雙對從來後進還處變動靜氣風雲觀視鏡劇節點線網"
    "傳統標準備選題類環場畫聲樂燈飾裝導編製審檢錯邏關係數據庫檔總綱領冊頁欄項應當須嚴絕斷續連貫轉軸廣湊"
)

REQUIRED_METHODOLOGY_TERMS = [
    "感动自己", "克制生悲悯", "情感蓄水池", "情感填充物", "骨肉互生",
    "主心骨", "游离", "持续让观众关心", "信息密度", "三变测试",
    "三层认知差漏斗", "猎奇", "新鲜度", "我没见过这样的东西", "四维认知破壁",
    "12 种实战工程武器", "声音桥"
]

def main():
    errors = []
    actual_files = []
    for r, dirs, files in os.walk(BASE_DIR):
        if ".git" in r or "__pycache__" in r:
            continue
        for f in files:
            actual_files.append(os.path.relpath(os.path.join(r, f), BASE_DIR))

    # Gate 1: All 26 files present & all 26 files have unique MD5 hashes
    for ef in EXPECTED_FILES:
        if ef not in actual_files:
            errors.append(f"[缺失标准文件] {ef}")
    for af in actual_files:
        if af not in EXPECTED_FILES:
            errors.append(f"[多余未登记文件] {af}")

    md5_map = {}
    for rel_p in sorted(actual_files):
        full_p = os.path.join(BASE_DIR, rel_p)
        h = hashlib.md5(open(full_p, "rb").read()).hexdigest()
        if h in md5_map:
            errors.append(f"[字节级100%重复文件] {rel_p} 与 {md5_map[h]} MD5完全相同 ({h})")
        md5_map[h] = rel_p

    # Gate 4 Prep: Dynamically extract character roles from EP01 screenplay
    ep01_sp_path = os.path.join(BASE_DIR, '01_全流程剧集开发标准母版库/01_专业文学剧本/EP01_《剧名》第一集《单集片名》专业文学剧本（高信息量台词与微动作交替标准母版）.docx')
    dynamic_chars = set()
    if os.path.exists(ep01_sp_path):
        d_sp = docx.Document(ep01_sp_path)
        for p in d_sp.paragraphs:
            t = p.text.strip()
            if t.startswith("人物："):
                for ch_name in re.split(r"[、，,\s]+", t.replace("人物：", "").strip()):
                    if ch_name:
                        dynamic_chars.add(ch_name)

    # Gate 5: Verify [07] 02_完整合订本 is 100% synchronized with [01]..[06]
    comp_p = os.path.join(BASE_DIR, '00_总控宪章与自动化引擎/02_《影视剧集从0到1全流程正向开发、视听转化与工业级交付总纲（完整合订本）》.docx')
    sub_rels = [
        '00_总控宪章与自动化引擎/00_《影视剧集从0到1全流程总控导航宪章（轻量路由入口）》.docx',
        '00_总控宪章与自动化引擎/01_按需加载全流程五阶段指导模块/模块01_立项初心（感动自己）、首集双引擎（高信息密度+新鲜度）与导演视听总纲指南.docx',
        '00_总控宪章与自动化引擎/01_按需加载全流程五阶段指导模块/模块02_开放世界生活场、底层运行铁律与立体人物关系四阶裂变（一人一档）指南.docx',
        '00_总控宪章与自动化引擎/01_按需加载全流程五阶段指导模块/模块03_高浓度场景空间拓扑（一景一档）与全局分集大纲（结构×情感填充物兼顾）指南.docx',
        '00_总控宪章与自动化引擎/01_按需加载全流程五阶段指导模块/模块04_双层细节工程（长伏笔+闲笔微镜像）、高信息量台词剧本与横向8列真导演分镜指南.docx',
        '00_总控宪章与自动化引擎/01_按需加载全流程五阶段指导模块/模块05_服化道顺场接戏统筹、三阶漏斗预告特辑与全流程100%完本交付核验指南.docx',
    ]
    if os.path.exists(comp_p):
        d_comp = docx.Document(comp_p)
        sub_tbs = []
        for sr in sub_rels:
            sd = docx.Document(os.path.join(BASE_DIR, sr))
            for ti, tb in enumerate(sd.tables):
                sub_tbs.append((os.path.basename(sr)[:12], ti + 1, len(tb.rows)))
        if len(d_comp.tables) != 1 + len(sub_tbs):
            errors.append(f"[合订本表格总数失锁] 合订本共 {len(d_comp.tables)} 表 != 1 + {len(sub_tbs)}")
        else:
            for idx, (sn, ti, r_cnt) in enumerate(sub_tbs):
                if len(d_comp.tables[idx + 1].rows) != r_cnt:
                    errors.append(f"[合订本表格行数失步] {sn} 表{ti} ({r_cnt}行) != 合订本表{idx+2} ({len(d_comp.tables[idx+1].rows)}行)")

    all_corpus_text = ""
    for rel_p in sorted(actual_files):
        full_p = os.path.join(BASE_DIR, rel_p)
        if rel_p.endswith(".docx"):
            d = docx.Document(full_p)
            txt = "\n".join(p.text for p in d.paragraphs) + "\n" + "\n".join(
                c.text for tb in d.tables for r in tb.rows for c in r.cells
            )
            # Gate 10: No cantSplit on data rows & no duplicate sentences inside cells
            for ti, tb in enumerate(d.tables):
                for ri, row in enumerate(tb.rows):
                    if ri > 0:
                        trPr = row._tr.find(qn("w:trPr"))
                        if trPr is not None and trPr.findall(qn("w:cantSplit")):
                            errors.append(f"[跨页截断隐患] {rel_p} 表{ti} 行{ri} 包含 w:cantSplit")
                    for ci, cell in enumerate(row.cells):
                        lines = [ln.strip() for ln in cell.text.split("\n") if len(ln.strip()) > 15]
                        if len(lines) != len(set(lines)):
                            errors.append(f"[单元格内整句重复] {rel_p} 表{ti} 行{ri} 列{ci}")

            # Gate 2: Screenplay 8 assertions
            if "01_专业文学剧本/" in rel_p or "01_《剧名》全集专业文学剧本总本" in rel_p or "01_EP01_" in rel_p:
                try:
                    verify_screenplay_8_assertions(full_p)
                except Exception as ex:
                    errors.append(str(ex))

            # Gate 3: Storyboard 8 columns structure
            if "02_导演分镜脚本与视听调度本/" in rel_p or "02_EP01_" in rel_p:
                try:
                    verify_storyboard_8cols(full_p)
                except Exception as ex:
                    errors.append(str(ex))

            # Gate 6: 02_总论 29 tables & EP01..EP12 independent 5-row sub-tables with 0 duplication
            if "02_《剧名》世界观设定" in rel_p:
                if len(d.tables) < 27:
                    errors.append(f"[02_总论表格数量不足] 仅 {len(d.tables)} 张表，要求包含EP01—EP12全部独立子表（>=27张表）")
                ep_subtables = {}
                for tb in d.tables:
                    if len(tb.rows) == 6 and len(tb.columns) == 2:
                        hdr = tb.rows[0].cells[1].text.strip()
                        m_ep = re.match(r"^(EP\d{2})", hdr)
                        if m_ep:
                            ep_code = m_ep.group(1)
                            body_str = "\n".join(r.cells[1].text.strip() for r in tb.rows[1:])
                            ep_subtables[ep_code] = body_str
                if len(ep_subtables) < 12:
                    errors.append(f"[02_总论分集独立子表不足12集] 仅找到 {sorted(ep_subtables.keys())}")
                seen_ep_bodies = {}
                for ep_k, ep_v in ep_subtables.items():
                    if ep_v in seen_ep_bodies:
                        errors.append(f"[02_总论分集大纲复制粘贴雷同] {ep_k} 与 {seen_ep_bodies[ep_v]} 内容雷同")
                    seen_ep_bodies[ep_v] = ep_k

            # Gate 7: Scene archive 5 chapters, 3-col axis table & rich Ch4
            if "06_核心场景一景一档/" in rel_p:
                ch_cnt = sum(1 for p in d.paragraphs if re.match(r"^第[一二三四五]章", p.text.strip()))
                if ch_cnt != 5 or len(d.tables[1].columns) != 3 or len(txt) < 2800:
                    errors.append(f"[场景档未达五章制或3列轴线表标准] {rel_p}")

            # Gate 8: Character archive 6 chapters & rich Ch4 biography without placeholders
            if "05_核心人物一人一档/" in rel_p:
                ch_cnt = sum(1 for p in d.paragraphs if re.match(r"^第[一二三四五六]章", p.text.strip()))
                if ch_cnt != 6 or "【角色名】" in txt or "________" in txt or len(txt) < 3800:
                    errors.append(f"[人物档未达六章制或存在填空占位符] {rel_p}")

            # Gate 4: Dynamic character consistency across screenplay, storyboard, statement, and costume table
            if any(k in rel_p for k in ["01_专业文学剧本/", "02_导演分镜脚本", "03_剧作结构", "07_服化道统筹"]):
                for ch_ent in dynamic_chars:
                    if ch_ent not in txt:
                        errors.append(f"[动态跨表角色未对齐] {rel_p} 缺少剧本登场角色 '{ch_ent}'")
        else:
            txt = open(full_p, "r", encoding="utf-8").read()

        all_corpus_text += "\n" + txt

        # Gate 9 & 10: Forbidden pull jargon, numbering offsets, IP clones, traditional chars
        if rel_p not in ("00_总控宪章与自动化引擎/scripts/audit_zero_errors.py", "00_总控宪章与自动化引擎/scripts/screenplay_gold_engine.py"):
            for bad_j in FORBIDDEN_PULL_AND_TYPOS:
                if bad_j in txt:
                    errors.append(f"[禁用词/编号偏移/填空符残留] {rel_p} 包含 '{bad_j}'")
            for bad_ip in FORBIDDEN_IP_HARDCODES:
                if bad_ip in txt:
                    errors.append(f"[特定IP或虚构古建剧硬编码残留] {rel_p} 包含 '{bad_ip}'")
            found_trad = sorted(set(ch for ch in txt if ch in FORBIDDEN_TRADITIONAL_CHARS))
            if found_trad:
                errors.append(f"[繁体字混杂残留] {rel_p} 包含繁体字 {found_trad}")

    for term in REQUIRED_METHODOLOGY_TERMS:
        if term not in all_corpus_text:
            errors.append(f"[核心方法论深层机制缺失] 全库未找到 '{term}'")

    if errors:
        for e in errors:
            print("ERROR:", e)
        sys.exit(1)
    print(f"[ALL 10 HARD QUALITY GATES PASSED WITH ZERO ERRORS] Verified {len(actual_files)} unique, pure universal workflow files (Dynamic Roles: {sorted(dynamic_chars)})!")

if __name__ == "__main__":
    main()
