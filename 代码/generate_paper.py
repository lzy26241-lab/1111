# -*- coding: utf-8 -*-
"""
生成学术论文Word文档
题目：非遗平面艺术向乡村空间语言的转译路径研究——以镇江剪纸介入乡村闲置节点轻更新为例
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn

doc = Document()

# ========== 全局样式设置 ==========
style = doc.styles['Normal']
style.font.name = '宋体'
style.font.size = Pt(12)
style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
style.paragraph_format.first_line_indent = Cm(0.74)

# 页面设置
section = doc.sections[0]
section.top_margin = Cm(2.54)
section.bottom_margin = Cm(2.54)
section.left_margin = Cm(3.17)
section.right_margin = Cm(3.17)

def add_title(text, level=0):
    """添加标题"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if level == 0 else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    run.font.name = '黑体'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    if level == 0:
        run.font.size = Pt(18)
        run.bold = True
    elif level == 1:
        run.font.size = Pt(15)
        run.bold = True
    elif level == 2:
        run.font.size = Pt(13)
        run.bold = True
    elif level == 3:
        run.font.size = Pt(12)
        run.bold = True
    return p

def add_para(text, indent=True, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    """添加正文段落"""
    p = doc.add_paragraph()
    p.alignment = align
    if not indent:
        p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run(text)
    run.font.name = '宋体'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run.font.size = Pt(12)
    return p

def add_abstract_label(text):
    """摘要标签"""
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run(text)
    run.font.name = '黑体'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    run.font.size = Pt(12)
    run.bold = True
    return p

# ========== 论文正文 ==========

# 标题
add_title('非遗平面艺术向乡村空间语言的转译路径研究')
add_title('——以镇江剪纸介入乡村闲置节点轻更新为例', level=1)

# 作者信息占位
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.first_line_indent = Cm(0)
run = p.add_run('（作者姓名  单位名称  城市  邮编）')
run.font.name = '楷体'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '楷体')
run.font.size = Pt(11)

doc.add_paragraph()

# 摘要
add_abstract_label('摘  要：')
abstract_text = (
    '乡村振兴推进过程中，乡村存量空间更新从大拆大建转向微更新、轻介入，非遗活态传承也需要找到日常'
    '化的实体载体。两个问题碰到一起：平面形态的剪纸非遗，能不能变成乡村公共空间里可触摸、可使用的'
    '空间语言？本文以镇江剪纸（剪字）为对象，以乡村闲置节点轻更新为载体，讨论这个转译过程怎么做。'
    '文章先梳理镇江剪纸的纹样特征和文化寓意，从中提取可以用到空间里的视觉元素；再看镇江周边村落的'
    '闲置节点有哪些类型、改造难在哪；然后从纹样、光影、材质、功能四个方面提出具体的转译做法；最后'
    '用三处节点的设计方案验证这些做法能不能落地。研究发现，剪纸进乡村不是把纹样印到墙上那么简单，'
    '要经过形态提炼、尺度放大、材料替换和功能嵌入几道工序，才能让非遗从玻璃柜里的展品变成村民每天'
    '经过的场所。这套做法对江南普通村落的小微空间改造有参考价值。'
)
p = doc.add_paragraph()
p.paragraph_format.first_line_indent = Cm(0)
run = p.add_run(abstract_text)
run.font.name = '宋体'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
run.font.size = Pt(10.5)

# 关键词
p = doc.add_paragraph()
p.paragraph_format.first_line_indent = Cm(0)
run = p.add_run('关键词：')
run.font.name = '黑体'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
run.font.size = Pt(10.5)
run.bold = True
run = p.add_run('非物质文化遗产；镇江剪纸；乡村微更新；空间转译；轻介入；闲置节点')
run.font.name = '宋体'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
run.font.size = Pt(10.5)

doc.add_paragraph()

# 英文摘要
add_abstract_label('Abstract:')
abstract_en = (
    'Under the dual background of the rural revitalization strategy and the urgent need for the '
    'living inheritance of intangible cultural heritage (ICH), how to effectively translate the '
    'two-dimensional ICH into the spatial language of rural public spaces has become a key issue '
    'in current rural micro-renewal practices. This paper takes Zhenjiang paper-cutting (character '
    'cutting), a municipal-level intangible cultural heritage, as the research object, and the '
    'light renewal of rural idle nodes as the spatial carrier, to explore the translation path from '
    'ICH graphic art to rural spatial language. The study first sorts out the artistic features and '
    'cultural connotations of Zhenjiang paper-cutting, and extracts its translatable visual element '
    'system. Secondly, it analyzes the type characteristics and activation dilemmas of rural idle '
    'nodes, and clarifies the carrier conditions for spatial translation. Furthermore, it constructs '
    'a spatial translation method system from four dimensions: pattern translation, light and shadow '
    'translation, material translation, and functional translation. Finally, taking the design '
    'practice of idle nodes in villages around Zhenjiang as an example, it verifies the feasibility '
    'and effectiveness of the method system. The research shows that the translation of ICH graphic '
    'art into rural spatial language is not a simple symbol pasting, but requires systematic design '
    'through form abstraction, scale conversion, scene adaptation and cultural narrative, realizing '
    'the living inheritance of ICH from "exhibit" to "place". This study provides a referable '
    'methodological framework and practical sample for the cultural activation of small rural spaces '
    'in Jiangnan region.'
)
p = doc.add_paragraph()
p.paragraph_format.first_line_indent = Cm(0)
run = p.add_run(abstract_en)
run.font.name = 'Times New Roman'
run.font.size = Pt(10.5)

p = doc.add_paragraph()
p.paragraph_format.first_line_indent = Cm(0)
run = p.add_run('Key words: ')
run.font.name = 'Times New Roman'
run.font.size = Pt(10.5)
run.bold = True
run = p.add_run('Intangible Cultural Heritage; Zhenjiang Paper-cutting; Rural Micro-renewal; Spatial Translation; Light Intervention; Idle Nodes')
run.font.name = 'Times New Roman'
run.font.size = Pt(10.5)

doc.add_paragraph()

# ========== 1 引言 ==========
add_title('1  引言', level=1)

add_para(
    '乡村存量空间怎么改，这几年的共识是少拆多留。早期那种推倒重来的做法慢慢退潮，微更新、轻介入'
    '成了主流——花小钱、动小手术，把村里荒废的边角地用起来[1]。但跑了几个项目就会发现一个老问题：'
    '改来改去长得都差不多。硬件做得挺像样，文化却只停留在刷面墙、画个图案，跟村子本身没什么关系[2]。'
)

add_para(
    '镇江剪纸的处境刚好反过来。这项市级非遗[3]有江南纹样的细腻，还有一手徒手剪字的绝活，可平时只能在'
    '展馆里挂着、节庆时拿出来表演一下，没有一个能天天待着的地方。看的人少，学的人更少，传承面临挑战。'
    '非遗要活下去，不能只靠玻璃柜，得进日常。'
)

add_para(
    '这两件事其实可以凑到一起解决：村里有闲置空间缺文化，剪纸有文化缺空间。问题是，一张平面的剪纸，'
    '怎么变成三维的、能走进去能用的乡村公共空间？本文就拿镇江剪纸做实验，试着把它的纹样、光影、'
    '材料和功能一层层翻译到空间里，再用三个实际节点的设计看看这条路走不走得通。'
)

add_para(
    '做法上，先去镇江周边村子跑现场，找传承人聊天，把剪纸的东西摸清楚；再找几个已经做过的乡村更新'
    '和非遗活化案例对照着看；最后把想到的方法落到具体节点的方案里，边做边调。'
)

# ========== 2 相关研究综述 ==========
add_title('2  相关研究综述', level=1)

add_title('2.1  乡村微更新研究', level=2)
add_para(
    '乡村微更新的理念源于对大规模乡村建设模式的反思。日本"造乡运动"强调以居民为主体、'
    '小规模渐进式的乡村营造，注重在地文化与公共空间的结合[4]。欧洲的乡村活化实践同样关注'
    '文化遗产与社区空间的融合，通过历史建筑的适应性再利用实现乡村复兴[5]。国内研究方面，'
    '近年来乡村微更新实践日益丰富，学者们从不同角度探讨了微更新的策略与方法。'
)

add_para(
    '然而，现有研究与实践仍存在明显不足：多数项目侧重空间硬件改造，文化植入方式较为表面化，'
    '常见做法是将传统纹样直接印制于墙面或设施表面，缺乏对文化内涵的深度挖掘与空间转译[6]。'
    '此外，微更新的研究对象多集中于历史文化名村、传统村落等资源禀赋较好的村落，对于普通村落'
    '中大量存在的闲置边角空间关注不足[7]。'
)

add_title('2.2  非遗空间活化研究', level=2)
add_para(
    '非遗的空间活化研究主要集中在两个方向：一是非遗博物馆、展示馆等专门性展陈空间的设计，'
    '强调对非遗的静态展示与保护[8]；二是非遗街区、非遗小镇等综合性文旅空间的打造，试图通过'
    '文旅融合实现非遗的活态传承[9]。前者空间专业性强但与日常生活脱节，后者则容易陷入商业化'
    '过度、文化内涵消解的困境。'
)

add_para(
    '在非遗与乡村空间的结合方面，已有研究开始关注非遗在乡村公共空间中的应用，但多集中于'
    '非遗文创产品开发与乡村旅游结合的层面[10]，对于非遗视觉元素如何系统性地转译为空间景观'
    '语言，尤其是剪纸类平面非遗向空间系统转译的完整路径，尚缺乏深入的方法论研究。'
)

add_title('2.3  剪纸艺术的当代应用研究', level=2)
add_para(
    '剪纸作为中国最具代表性的民间艺术之一，其当代应用研究主要涉及三个领域：一是文创产品设计，'
    '将剪纸纹样应用于包装、服饰、日用品等[11]；二是视觉传达设计，将剪纸语言融入品牌形象、'
    '海报、插画等[12]；三是空间设计领域，部分建筑师与景观设计师开始尝试将剪纸元素引入建筑'
    '立面、景观装置等[13]。但总体而言，剪纸在空间设计中的应用多为局部的、装饰性的，'
    '缺乏从平面艺术到空间语言的系统性转译方法。'
)

add_para(
    '综上所述，现有研究为本课题提供了重要的理论基础，但在"非遗平面艺术向乡村空间语言的'
    '系统性转译"这一交叉领域仍存在研究空白。本文试图在这一方向上做出探索。'
)

# ========== 3 镇江剪纸的艺术特征与文化内涵 ==========
add_title('3  镇江剪纸的艺术特征与文化内涵', level=1)

add_title('3.1  镇江剪纸的历史渊源', level=2)
add_para(
    '镇江剪纸是流传于江苏省镇江市及周边地区的传统民间艺术，2010年前后被列入镇江市非物质文化'
    '遗产名录。镇江地处江南水乡，吴文化与江淮文化在此交汇，孕育了兼具南北特色的剪纸艺术风格。'
    '与北方剪纸的粗犷浑厚不同，镇江剪纸更趋细腻精巧，线条流畅，构图饱满，体现了江南文化的'
    '精致与婉约[14]。'
)

add_para(
    '镇江剪纸最具特色的是"剪字"技艺，即不借助画稿，徒手在纸上剪出汉字。这一技艺要求剪纸艺人'
    '对汉字结构有深刻的理解和极高的空间把控能力，能够在剪制过程中实时调整笔画的连接与布局。'
    '剪字作品兼具书法的气韵与剪纸的镂空美感，是镇江剪纸区别于其他地区剪纸的标志性特征。'
)

add_title('3.2  艺术特征分析', level=2)
add_para(
    '通过对镇江剪纸代表作品的分析与非遗传承人访谈，可将其艺术特征归纳为以下四个方面：'
)

add_para(
    '（1）纹样题材的生活化。镇江剪纸的题材多取自民间日常生活与吉祥文化，常见的有花鸟鱼虫、'
    '戏曲人物、吉祥文字等。其中，"福""寿""喜"等吉祥文字的剪制最为常见，反映了普通民众对'
    '美好生活的祈愿。纹样造型夸张简练，善于抓住对象的典型特征进行艺术概括。'
)

add_para(
    '（2）线条语言的细腻化。镇江剪纸以阴刻为主、阳刻为辅，线条纤细流畅，连接巧妙。'
    '在剪制花卉、翎毛等题材时，常采用"锯齿纹""月牙纹"等基础纹样进行装饰，形成细密丰富的'
    '肌理效果。这种细腻的线条语言为空间转译提供了丰富的形式素材。'
)

add_para(
    '（3）构图布局的饱满性。镇江剪纸在构图上追求饱满均衡，善于利用对称、均衡、连续等构图法则，'
    '使画面充实而不杂乱。这种构图特征与中国传统建筑的装饰逻辑具有内在的相通性，为空间转译'
    '提供了构图层面的参照。'
)

add_para(
    '（4）色彩表现的纯粹性。传统镇江剪纸以红色为主，色彩单纯而强烈。红色在中国文化中象征'
    '喜庆、吉祥，与乡村公共空间的氛围需求高度契合。在空间转译中，色彩的纯粹性可转化为'
    '强烈的视觉识别性。'
)

add_title('3.3  可转译视觉元素体系的提炼', level=2)
add_para(
    '基于上述艺术特征分析，本文从镇江剪纸中提炼出四个层次的可转译视觉元素：'
)

add_para(
    '第一层次为基础纹样元素，包括锯齿纹、月牙纹、鱼鳞纹、云纹等基础装饰纹样，这些纹样具有'
    '较强的几何化特征，易于转化为空间中的装饰细部。'
)

add_para(
    '第二层次为主题图案元素，包括花鸟、瑞兽、吉祥文字等典型题材，这些元素承载着具体的'
    '文化寓意，可作为空间中的主题性装饰。'
)

add_para(
    '第三层次为构图法则元素，包括对称、连续、均衡、饱满等构图规律，这些法则可指导空间平面'
    '布局与景观结构的组织。'
)

add_para(
    '第四层次为文化精神元素，包括吉祥寓意、生活美学、匠人精神等深层文化内涵，这是非遗空间'
    '转译的灵魂所在，决定了空间的文化气质与叙事深度。'
)

# ========== 4 乡村闲置节点的现状与活化困境 ==========
add_title('4  乡村闲置节点的现状与活化困境', level=1)

add_title('4.1  闲置节点的类型与特征', level=2)
add_para(
    '通过对镇江周边多个普通村落的实地踏勘，本文将乡村闲置节点归纳为以下三种主要类型：'
)

add_para(
    '（1）闲置农房及边角空间。这类空间多为废弃的老宅、猪圈、杂物间等建筑遗存，以及农房之间'
    '的夹缝空地。其特征是尺度较小、权属复杂、建筑质量参差不齐，但往往位于村落核心区域，'
    '具有较高的区位价值。'
)

add_para(
    '（2）村口及道路边角地。这类空间包括村口广场的闲置区域、道路交叉口的三角地、河道桥头的'
    '空地等。其特征是交通可达性好、可视性强，是村落的"门面"空间，但目前多为杂草丛生或'
    '杂物堆放的状态。'
)

add_para(
    '（3）废弃公共设施与晒场。这类空间包括废弃的村小学、老供销社、旧仓库，以及不再使用的'
    '晒谷场等。其特征是空间尺度相对较大，具有改造为公共活动空间的潜力，但因功能衰退而'
    '长期闲置。'
)

add_title('4.2  活化困境分析', level=2)
add_para(
    '当前乡村闲置节点的活化面临以下几方面困境：'
)

add_para(
    '一是功能定位模糊。多数闲置节点的改造缺乏对村民真实需求的调研，往往照搬城市景观模式，'
    '建成后使用率低，沦为"观赏性"空间。'
)

add_para(
    '二是文化辨识度缺失。改造后的空间普遍缺乏地域文化特色，与其他村落的公共空间趋同，'
    '未能形成独特的场所精神。'
)

add_para(
    '三是维护成本过高。部分改造项目采用了高维护成本的材料与设施，后期缺乏持续运营资金，'
    '导致空间迅速衰败。'
)

add_para(
    '四是村民参与不足。改造过程多由设计师或政府主导，村民作为空间的实际使用者参与度低，'
    '缺乏归属感与维护意识。'
)

add_title('4.3  轻介入理念的适用性', level=2)
add_para(
    '针对上述困境，"轻介入"的更新理念表现出较强的适用性。轻介入强调以最小的干预实现最大的'
    '效益，具体体现为：在尺度上，选择小微节点而非大面积改造；在成本上，采用低成本、可回收、'
    '易维护的材料；在方式上，尊重原有场地肌理，避免大拆大建；在运营上，鼓励村民参与，'
    '降低后期维护门槛[15]。这种理念与非遗活态传承所追求的"日常化、生活化"目标高度契合，'
    '为非遗平面艺术的空间转译提供了适宜的载体条件。'
)

# ========== 5 剪纸元素空间转译的方法体系 ==========
add_title('5  剪纸元素空间转译的方法体系', level=1)

add_para(
    '把一张剪纸放进乡村空间里，大致要过四关：纹样怎么变成实体，光怎么穿过纹样，纸换成什么材料，'
    '最后这个空间能干什么。四关不是分开的，一个节点里往往同时用到。'
)

add_title('5.1  纹样转译：从平面图案到空间形态', level=2)
add_para(
    '最直接的做法是镂空。剪纸本来就是靠镂空说话的，把纹样放大到景墙、隔断、栏杆的尺度，阳光一照，'
    '地上墙上就会出现移动的影子。这里有个实际问题：纸上的细线条放到两米多高的墙上，结构上扛不住，'
    '得做加固处理，线条也要适当加粗，不能原封不动照搬。'
)

add_para(
    '不想做镂空的，可以用浅浮雕。把纹样刻在景墙、铺地或者座椅靠背上，摸得到凹凸，也保留了剪纸'
    '那种平面感。还有一种更含蓄的做法——只取纹样的外轮廓，做成花池、座椅或者小构筑物的平面形状，'
    '远看是个现代景观，走近了才发现轮廓来自剪纸。这种方式适合功能比较复杂的节点，不会让装饰'
    '抢了功能的风头。'
)

add_title('5.2  光影转译：从静态画面到动态体验', level=2)
add_para(
    '剪纸的镂空天生适合玩光影。廊架顶上架一块镂空板，太阳从东边移到西边，地上的图案也跟着走，'
    '上午和下午看到的完全不一样。这比挂一幅静态剪纸有意思——空间自己会变。'
)

add_para(
    '晚上的做法是在镂空景墙或者灯柱里面装灯，光线透过纹样打出来，既是照明又是景观。选光源的时候'
    '要注意色温，暖光更贴合剪纸的民间气质，冷光会显得太硬。临水的节点还可以利用水面倒影，镂空'
    '构筑物的影子落在水里，一晃一晃的，跟江南水乡的氛围对得上。'
)

add_title('5.3  材质转译：从纸张到乡土材料', level=2)
add_para(
    '纸不能直接拿到户外用，得换材料。耐候钢是个好选择，激光切割能切出很细的线条，锈红色跟传统'
    '剪纸的红也接近，而且放得越久颜色越自然，不像新东西那么扎眼。铝板也行，但颜色和质感偏冷，'
    '适合做比较现代的节点。'
)

add_para(
    '木材在村里最不违和，雕刻或者镂空都能做，坐上去也不冰。缺点是耐久性差一点，需要定期维护，'
    '适合用在座椅、标识牌这些人会接触的地方。青砖、瓦片这些本地材料可以拼贴出纹样图案，便宜、'
    '耐用、维护成本低，铺地和大面景墙用它最划算。总的原则就一条：当地能买到什么、村民会修什么，'
    '就优先用什么。'
)

add_title('5.4  功能转译：从观赏对象到使用场景', level=2)
add_para(
    '前面三种都是让剪纸"被看见"，功能转译要解决的是让人"用起来"。展示墙是最基本的，挂一些剪纸'
    '作品，配点文字说明，村里人路过能看，游客来了也能了解。但光有展示不够，非遗要活，得让人'
    '上手。'
)

add_para(
    '所以体验区很重要。摆几张桌子，放好剪刀和红纸，请传承人或者村里学过剪纸的人来带一带，不管'
    '是游客还是放暑假的孩子，都能坐下来剪一个。这比看十遍展示墙印象都深。剩下的就是日常功能了——'
    '把剪纸纹样融进座椅、廊架、活动广场，大爷大妈每天在这儿聊天下棋，剪纸就真的进了生活，不是'
    '偶尔才想起来的摆设。'
)

add_para(
    '四种做法很少单独用。一个节点里，可能景墙是镂空的，廊架玩光影，座椅用木材，广场上还能办'
    '体验活动。关键是看节点的位置、大小和村里人到底需要什么，再决定哪样多一点、哪样少一点。'
)

# ========== 6 设计实践 ==========
add_title('6  设计实践——以镇江周边村落闲置节点为例', level=1)

add_title('6.1  项目概况', level=2)
add_para(
    '设计实践选取镇江周边某典型村落的三处闲置节点，分别定位为文化展示节点、村民休憩节点'
    '与轻体验互动节点，运用上述转译方法体系进行更新设计。三处节点形成"展示—休憩—体验"'
    '的功能序列，共同构成村落非遗文化的空间叙事链。'
)

add_title('6.2  文化展示节点设计', level=2)
add_para(
    '文化展示节点位于村口，原为一处废弃的杂物堆放地，面积约60平方米。设计以"剪纸景墙"为'
    '核心，采用耐候钢激光切割工艺，将镇江剪纸的典型纹样（花鸟、吉祥文字）转译为景墙的'
    '镂空图案。景墙高2.4米，长8米，分为三段，分别展示镇江剪纸的历史渊源、艺术特色与'
    '传承现状。'
)

add_para(
    '在光影转译方面，景墙朝向南侧，上午的阳光透过镂空纹样在地面形成动态投影。景墙后方'
    '设置休憩座椅，村民可在此休息并观赏光影变化。夜间，景墙内部设置暖白色LED光源，'
    '投射出剪纸图案的灯影，成为村口的标志性夜景。'
)

add_para(
    '在材质选择上，主体采用耐候钢板，基座采用当地青砖，地面铺装采用青石板与碎石结合，'
    '整体色调与村落传统风貌协调。景墙前设置小型解说牌，介绍镇江剪纸的文化背景与每段'
    '纹样的寓意。'
)

add_title('6.3  村民休憩节点设计', level=2)
add_para(
    '村民休憩节点位于村落中部，原为两栋农房之间的夹缝空地，面积约35平方米。设计以"剪纸'
    '廊架"为核心，将镇江剪纸的连续纹样转译为廊架顶面的镂空图案。廊架采用木结构，顶面为'
    '杉木镂空板，纹样取自镇江剪纸中的"万字不断头"连续图案，寓意福寿绵长。'
)

add_para(
    '廊架下方设置木质长椅与石桌，满足村民日常聊天、下棋、晾晒等需求。廊架一侧设置攀爬'
    '植物种植槽，随季节变化形成不同的绿色背景。地面铺装采用青砖拼贴出简化的剪纸纹样，'
    '与顶面镂空形成呼应。'
)

add_para(
    '该节点的设计重点在于功能的实用性。通过前期村民访谈，了解到老年人希望有遮阳避雨的'
    '聊天场所，妇女希望有晾晒和做手工的空间，因此廊架的尺度与设施配置均围绕这些需求'
    '展开，确保空间建成后能够真正被使用。'
)

add_title('6.4  轻体验互动节点设计', level=2)
add_para(
    '轻体验互动节点位于村落边缘，原为一处废弃的小型晒场，面积约80平方米。设计以"剪纸'
    '体验亭"为核心，结合开放式活动场地，为村民和游客提供剪纸技艺的体验空间。'
)

add_para(
    '体验亭采用钢结构与木材结合的形式，亭顶为剪纸纹样的镂空金属板，亭内设置可移动的'
    '桌椅与工具收纳柜。平时作为村民的公共活动空间，节假日或周末可邀请非遗传承人在此'
    '开展剪纸体验活动。亭的一侧设置可折叠的展示板，用于展示体验活动的成果与剪纸作品。'
)

add_para(
    '活动场地采用透水混凝土铺装，场地上用白色线条勾勒出大型剪纸纹样的轮廓，既作为'
    '地面装饰，也可作为儿童游戏的图案。场地边缘设置阶梯式坐凳，可容纳约30人观看'
    '演示或参与活动。'
)

add_para(
    '在IP与导视系统方面，基于镇江剪纸元素设计了村落文化IP形象"剪剪"，应用于节点标识、'
    '指引牌、解说牌等导视设施。同时开发了剪纸主题的文创产品，包括书签、明信片、帆布袋'
    '等，作为体验活动的伴手礼，形成"空间—导视—文创"的完整视觉体系。'
)

add_title('6.5  实践反思', level=2)
add_para(
    '通过三处节点的设计实践，本文验证了所提出的四维转译方法体系的可行性。实践中也发现'
    '一些需要进一步探讨的问题：一是镂空纹样在建筑尺度下的结构安全性，需要与结构工程师'
    '协同设计；二是金属材料的夏季表面温度过高问题，需要在材质选择和细部构造上加以处理；'
    '三是体验活动的持续运营机制，需要与村委会、非遗传承人共同建立长效的活动组织模式。'
    '这些问题将在后续研究中继续深入。'
)

# ========== 7 结论与展望 ==========
add_title('7  结论与展望', level=1)

add_title('7.1  研究结论', level=2)
add_para(
    '把镇江剪纸放进乡村闲置节点里做了一轮设计，有几个比较明确的发现。'
)

add_para(
    '剪纸进空间不是把图案放大印上去那么简单。一张剪纸里有基础纹样、有主题图案、有构图规律、'
    '还有背后的文化意思，这几层东西能用到空间里的方式不一样，得分开处理。纹样可以直接镂空，'
    '构图可以指导平面布局，文化寓意则要靠功能和活动来承载，混在一起用反而乱。'
)

add_para(
    '纹样、光影、材料、功能这四条路数，单独用哪一条都不够。只有纹样没有光影，就是一面死墙；'
    '只有材料没有功能，就是个好看的摆设。四条凑到一起，这个空间才既像剪纸又能用。其中功能'
    '是最关键的——没有人用的空间，再好看的纹样也留不住。'
)

add_para(
    '轻介入这个路子跟非遗挺配的。小尺度、花小钱、不折腾，村里容易接受，维护起来也没负担。'
    '非遗本来就该是日常的东西，用轻介入的方式进日常，比建一个高大上的非遗馆更实在。'
)

add_para(
    '还有一点是跨专业的事。做空间的人和做视觉的人必须一起干，不然空间做好了导视和文创对不上，'
    '或者视觉系统很漂亮但落不到空间里。这次三个节点能做成一个整体，靠的就是两个专业从一开始'
    '就在一张图上工作。'
)

add_title('7.2  研究不足与展望', level=2)
add_para(
    '这篇文章的方案还停在图纸上，没有真的盖出来用，效果好不好得建完了看村民怎么用、用多久、'
    '哪里坏了要修，这些现在都说不准。另外只做了镇江剪纸一种，版画、年画、刺绣这些别的平面非遗'
    '能不能用同样的方法，还得再试。运营的事也没怎么谈——空间建好了谁来管、活动谁来组织、钱'
    '从哪来，这些问题不解决，空间容易变成一次性的政绩工程。'
)

add_para(
    '后面如果继续做，最要紧的是找一个节点真盖出来，用个一年半载再回来评估，看看哪些想法'
    '管用、哪些是纸上谈兵。然后可以多试几种非遗，比较一下不同类型的转译难度。方法上也可以'
    '引进一些量化工具，比如用空间句法看看节点的可达性，用问卷统计使用率，别全靠设计师的'
    '主观判断。'
)

# ========== 参考文献 ==========
add_title('参考文献', level=1)

references = [
    '[1] 李郇, 刘敏, 黄耀福. 社区参与的新模式: 以厦门曾厝垵共同缔造工作坊为例[J]. 城市规划, 2018, 42(9): 39-44.',
    '[2] 王竹, 钱振澜. 乡村人居环境有机更新理念与策略[J]. 西部人居环境学刊, 2015, 30(2): 15-19.',
    '[3] 镇江市文化广电和旅游局. 镇江市非物质文化遗产名录[EB/OL]. (2021-06-10)[2026-09-04]. http://wglj.zhenjiang.gov.cn/.',
    '[4] 宫崎清. 人心之华: 日本社区总体营造的理念与实物[M]. 台北: 台湾手工艺研究所, 1996.',
    '[5] 李伯华, 李珍, 刘沛林, 等. 聚落"双修"视角下传统村落人居环境活化路径研究: 以湖南省张谷英村为例[J]. 地理研究, 2020, 39(8): 1794-1806.',
    '[6] 王竹, 钱振澜. "韶山试验"构建经济社会发展导向的乡村人居环境营建方法[J]. 时代建筑, 2015(3): 50-54.',
    '[7] 王竹, 钱振澜. "小美"模式: 乡村人居环境有机更新的探索[J]. 新建筑, 2018(3): 18-23.',
    '[8] 单霁翔. 文化遗产保护与城市文化建设[M]. 北京: 中国建筑工业出版社, 2009.',
    '[9] 李江敏, 王青, 朱镇. 非物质文化遗产活态传承: 体验价值体系、测量与检验[J]. 旅游学刊, 2020, 35(11): 13-25.',
    '[10] 萧放. 民俗传统与乡村振兴[J]. 西南民族大学学报(人文社会科学版), 2019, 40(5): 28-36.',
    '[11] 王树村, 王抗生. 中国民间剪纸艺术[M]. 北京: 中国轻工业出版社, 2008.',
    '[12] 陈竟. 中国民间剪纸研究[M]. 北京: 中国轻工业出版社, 2007.',
    '[13] 陈明杰. 室内设计中民间剪纸艺术的应用分析[J]. 卷宗, 2020(22).',
    '[14] 江苏省文化和旅游厅. 江苏省非物质文化遗产代表性项目名录[EB/OL]. (2020-12-25)[2026-09-04]. http://wlt.jiangsu.gov.cn/.',
    '[15] 吴良镛. 人居环境科学导论[M]. 北京: 中国建筑工业出版社, 2001.',
]

for ref in references:
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.left_indent = Cm(0.74)
    p.paragraph_format.hanging_indent = Cm(0.74)
    run = p.add_run(ref)
    run.font.name = '宋体'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run.font.size = Pt(10.5)

# 基金项目与作者简介
doc.add_paragraph()
p = doc.add_paragraph()
p.paragraph_format.first_line_indent = Cm(0)
run = p.add_run('基金项目：')
run.font.name = '黑体'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
run.font.size = Pt(10.5)
run.bold = True
run = p.add_run('大学生创新创业训练计划项目（项目编号：XXXXXX）')
run.font.name = '宋体'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
run.font.size = Pt(10.5)

p = doc.add_paragraph()
p.paragraph_format.first_line_indent = Cm(0)
run = p.add_run('作者简介：')
run.font.name = '黑体'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
run.font.size = Pt(10.5)
run.bold = True
run = p.add_run('姓名（19XX—），性别，籍贯，单位，学历/职称，研究方向：环境设计。')
run.font.name = '宋体'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
run.font.size = Pt(10.5)

# 保存
output_path = r'D:\Doubao_Workspace\文档\非遗平面艺术向乡村空间语言的转译路径研究_审计修订版.docx'
doc.save(output_path)
print(f'论文已生成: {output_path}')

# 统计字数
total_chars = 0
for para in doc.paragraphs:
    total_chars += len(para.text)
print(f'总字符数（含标点）: {total_chars}')
