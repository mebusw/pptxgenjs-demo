import pptxgen from "pptxgenjs";

// 1. Create a Presentation
let pres = new pptxgen();

// Metadata
pres.title = 'Joyson Factory Digital Transformation';
pres.author = 'Digital Transformation Team';
pres.subject = 'TOC Real-time Architecture';
pres.company = 'Joyson';

// Layout - Instruction: Use 16x9
pres.layout = 'LAYOUT_16x9';

// Define a clean Slide Master (Optional, but good for consistency)
pres.defineSlideMaster({
    title: "CLEAN_MASTER",
    background: { color: "F5F7FA" },
    slideNumber: { x: "95%", y: "92%", fontSize: 10, color: "999999" },
});

/**
 * Builds the specific slides for the presentation based on the Story Line.
 * @param {pptxgen} pres 
 */
let buildSlides = (pres) => {
    // --- Constants & Styles ---
    const COLORS = {
        PRIMARY: "003366",   // Deep Blue
        SECONDARY: "007ACC", // Bright Blue
        ACCENT: "FF8C00",    // Orange
        TEXT_MAIN: "333333",
        TEXT_LIGHT: "767676",
        BG_LIGHT: "F5F7FA",
        WHITE: "FFFFFF",
        LIGHT_GRAY: "E0E0E0",
        CHART_BAR: "003366"
    };

    const FONTS = {
        TITLE: "Microsoft YaHei",
        BODY: "Microsoft YaHei"
    };

    // Helper to create standard bullet options
    const bulletOpts = { code: "2022", indent: 20 }; // Bullet styling

    // ==========================================================================
    // Slide 1: Title Slide
    // ==========================================================================
    {
        let slide = pres.addSlide(); // No master for title to allow full bleed
        
        // Background Image (Abstract Tech)
        slide.addImage({
            path: "https://images.unsplash.com/photo-1518770660439-4636190af475?q=80&w=1920&auto=format&fit=crop",
            x: 0, y: 0, w: "100%", h: "100%",
            sizing: { type: "cover" }
        });
        
        // Semi-transparent overlay for readability
        slide.addShape(pres.ShapeType.rect, {
            x: 0, y: 0, w: "100%", h: "100%",
            fill: { color: "000000", transparency: 30 }
        });

        // Title
        slide.addText("均胜工厂数字化转型与TOC实时监控架构", {
            x: 0.5, y: 2, w: 9, h: 1.5,
            fontSize: 40, color: COLORS.WHITE, bold: true, align: "center", fontFace: FONTS.TITLE
        });

        // Subtitle
        slide.addText("从数据孤岛到基于MQTT的实时智能决策系统", {
            x: 1, y: 3.5, w: 8, h: 0.5,
            fontSize: 22, color: "EEEEEE", align: "center", fontFace: FONTS.BODY
        });

        // Presenter Info
        slide.addText("汇报人：数字化转型团队 | 时间：2025年11月", {
            x: 1, y: 5, w: 8, h: 0.5,
            fontSize: 14, color: "DDDDDD", align: "center", fontFace: FONTS.BODY
        });

        // Visual Prompt Note
        slide.addText("Visuals Prompt: Abstract digital network connecting factory machines, blue and white theme, high tech style.", {
            x: 0.2, y: 5.3, w: 8, h: 0.3, fontSize: 8, color: "AAAAAA"
        });
    }

    // ==========================================================================
    // Slide 2: Background & Objectives
    // ==========================================================================
    {
        let slide = pres.addSlide({ masterName: "CLEAN_MASTER" });

        // Title
        slide.addText("项目背景：COO的核心诉求与现状痛点", {
            x: 0.5, y: 0.3, w: 9, h: 0.5,
            fontSize: 28, color: COLORS.PRIMARY, bold: true, fontFace: FONTS.TITLE
        });

        // Left Column: Pain Points
        slide.addShape(pres.ShapeType.rect, { x: 0.5, y: 1.2, w: 4.2, h: 3.8, fill: COLORS.WHITE, line: { color: COLORS.LIGHT_GRAY, width: 1 } });
        slide.addText("当前痛点 (Current Pain Points)", { x: 0.7, y: 1.4, w: 3.8, h: 0.3, fontSize: 16, bold: true, color: COLORS.ACCENT, fontFace: FONTS.BODY });
        
        slide.addText([
            { text: "数据滞后：依赖月度财务报表，数据严重滞后，无法指导实时生产。", options: { breakLine: true, bullet: bulletOpts } },
            { text: "系统孤岛：SAP, HR, WMS各管一摊，存在大量手工Excel台账。", options: { breakLine: true, bullet: bulletOpts } },
            { text: "基础薄弱：缺乏统一的底层数据总线，无法支撑AI分析。", options: { bullet: bulletOpts } }
        ], { x: 0.7, y: 1.8, w: 3.8, h: 3, valign: "top", fontSize: 12, color: COLORS.TEXT_MAIN, fontFace: FONTS.BODY, paraSpaceAfter: 10 });

        // Right Column: Objectives
        slide.addShape(pres.ShapeType.rect, { x: 5.3, y: 1.2, w: 4.2, h: 3.8, fill: COLORS.WHITE, line: { color: COLORS.LIGHT_GRAY, width: 1 } });
        slide.addText("核心目标 (Core Objectives)", { x: 5.5, y: 1.4, w: 3.8, h: 0.3, fontSize: 16, bold: true, color: COLORS.PRIMARY, fontFace: FONTS.BODY });

        slide.addText([
            { text: "TOC实时化：实现Total Operation Cost的实时可视化与管理。", options: { breakLine: true, bullet: bulletOpts } },
            { text: "数据驱动：将运营KPI缩短至秒级更新，辅助管理层快速决策。", options: { breakLine: true, bullet: bulletOpts } },
            { text: "架构统一：建立标准化的实时数据总线。", options: { bullet: bulletOpts } }
        ], { x: 5.5, y: 1.8, w: 3.8, h: 3, valign: "top", fontSize: 12, color: COLORS.TEXT_MAIN, fontFace: FONTS.BODY, paraSpaceAfter: 10 });

        // Prompt Note
        slide.addText("Visuals Prompt: Comparison illustration: A stack of paper reports vs a futuristic glowing real-time dashboard on a tablet.", {
            x: 0.2, y: 5.3, w: 8, h: 0.3, fontSize: 8, color: "BBBBBB"
        });
    }

    // ==========================================================================
    // Slide 3: Architecture
    // ==========================================================================
    {
        let slide = pres.addSlide({ masterName: "CLEAN_MASTER" });

        slide.addText("技术架构转型：MQTT + Edge Computing", {
            x: 0.5, y: 0.3, w: 9, h: 0.5,
            fontSize: 28, color: COLORS.PRIMARY, bold: true, fontFace: FONTS.TITLE
        });

        // Diagram Construction
        // Center: MQTT Broker
        slide.addShape(pres.ShapeType.ellipse, { x: 4, y: 1.8, w: 2, h: 1.5, fill: COLORS.PRIMARY });
        slide.addText("MQTT Broker\n(Publish/Subscribe)", { x: 4, y: 1.8, w: 2, h: 1.5, align: "center", color: COLORS.WHITE, fontSize: 14, bold: true, fontFace: FONTS.BODY });

        // Left: Edge/PLC
        slide.addShape(pres.ShapeType.roundRect, { x: 0.8, y: 2.05, w: 2, h: 1, fill: COLORS.LIGHT_GRAY });
        slide.addText("Edge / PLC\n(Source Data)", { x: 0.8, y: 2.05, w: 2, h: 1, align: "center", fontSize: 12, color: COLORS.TEXT_MAIN, fontFace: FONTS.BODY });
        // Arrow L->C
        slide.addShape(pres.ShapeType.rightArrow, { x: 2.9, y: 2.4, w: 1, h: 0.3, fill: COLORS.SECONDARY });

        // Right Top: Dashboard
        slide.addShape(pres.ShapeType.roundRect, { x: 7.2, y: 1.3, w: 2, h: 0.8, fill: COLORS.LIGHT_GRAY });
        slide.addText("Dashboard\n(Real-time View)", { x: 7.2, y: 1.3, w: 2, h: 0.8, align: "center", fontSize: 12, color: COLORS.TEXT_MAIN, fontFace: FONTS.BODY });
        // Arrow C->R Top
        slide.addShape(pres.ShapeType.rightArrow, { x: 6.1, y: 1.6, w: 1, h: 0.2, fill: COLORS.SECONDARY, rotate: -20 });

        // Right Bottom: History DB
        slide.addShape(pres.ShapeType.roundRect, { x: 7.2, y: 3.0, w: 2, h: 0.8, fill: COLORS.LIGHT_GRAY });
        slide.addText("History DB\n(Analysis)", { x: 7.2, y: 3.0, w: 2, h: 0.8, align: "center", fontSize: 12, color: COLORS.TEXT_MAIN, fontFace: FONTS.BODY });
        // Arrow C->R Bottom
        slide.addShape(pres.ShapeType.rightArrow, { x: 6.1, y: 3.3, w: 1, h: 0.2, fill: COLORS.SECONDARY, rotate: 20 });

        // Text explanation at bottom
        slide.addText("架构优势与机制：", { x: 0.5, y: 4.0, w: 9, h: 0.3, fontSize: 14, bold: true, color: COLORS.PRIMARY, fontFace: FONTS.BODY });
        slide.addText([
            { text: "边缘驱动：工位产生原始数据，Edge端订阅并清洗。", options: { breakLine: true, bullet: bulletOpts } },
            { text: "轻量高效：摒弃传统轮询(Polling)，数据变化即推送，节省带宽。", options: { breakLine: true, bullet: bulletOpts } },
            { text: "层级结构：Enterprise -> Factory -> Line -> Station -> Topic。", options: { bullet: bulletOpts } }
        ], { x: 0.5, y: 4.3, w: 9, h: 1.2, fontSize: 12, color: COLORS.TEXT_MAIN, fontFace: FONTS.BODY, paraSpaceAfter: 5 });

        // Prompt Note
        slide.addText("Visuals Prompt: Technical architecture diagram. Center: MQTT Broker. Left: PLC/Sensors. Right: Dashboard/DB. Arrows showing flow.", {
            x: 0.2, y: 5.3, w: 8, h: 0.3, fontSize: 8, color: "BBBBBB"
        });
    }

    // ==========================================================================
    // Slide 4: Core Concept (Manifest + MongoDB)
    // ==========================================================================
    {
        let slide = pres.addSlide({ masterName: "CLEAN_MASTER" });

        slide.addText("核心配置逻辑：Manifest与MongoDB", {
            x: 0.5, y: 0.3, w: 9, h: 0.5,
            fontSize: 28, color: COLORS.PRIMARY, bold: true, fontFace: FONTS.TITLE
        });

        // Left: Content
        slide.addText("Manifest (工艺清单)", { x: 0.5, y: 1.2, w: 4, h: 0.4, fontSize: 18, bold: true, color: COLORS.SECONDARY, fontFace: FONTS.BODY });
        slide.addText([
            { text: "定义产品工艺路线、工位逻辑及防错规则。", options: { breakLine: true, bullet: bulletOpts } },
            { text: "配置下发至产线，驱动生产流程。", options: { bullet: bulletOpts } }
        ], { x: 0.5, y: 1.6, w: 4.5, h: 1.5, fontSize: 14, color: COLORS.TEXT_MAIN, fontFace: FONTS.BODY, paraSpaceAfter: 10 });

        slide.addText("MongoDB (非结构化存储)", { x: 0.5, y: 3.2, w: 4, h: 0.4, fontSize: 18, bold: true, color: COLORS.SECONDARY, fontFace: FONTS.BODY });
        slide.addText([
            { text: "适应工业数据结构多变需求（如随时增加新传感器字段）。", options: { breakLine: true, bullet: bulletOpts } },
            { text: "OPS模式：增加字段不影响旧数据。", options: { bullet: bulletOpts } }
        ], { x: 0.5, y: 3.6, w: 4.5, h: 1.5, fontSize: 14, color: COLORS.TEXT_MAIN, fontFace: FONTS.BODY, paraSpaceAfter: 10 });

        // Right: Image Placeholder for Workflow
        slide.addImage({
            path: "https://images.unsplash.com/photo-1581091226825-a6a2a5aee158?q=80&w=1000&auto=format&fit=crop",
            x: 5.5, y: 1.5, w: 4, h: 3.5,
            sizing: { type: "contain" }
        });

        // Prompt Note
        slide.addText("Visuals Prompt: A JSON document icon labeled 'Manifest' transforming into physical production steps on a conveyor belt.", {
            x: 0.2, y: 5.3, w: 8, h: 0.3, fontSize: 8, color: "BBBBBB"
        });
    }

    // ==========================================================================
    // Slide 5: TOC Data Breakdown - Variable Costs
    // ==========================================================================
    {
        let slide = pres.addSlide({ masterName: "CLEAN_MASTER" });

        slide.addText("TOC数据拆解：变动成本 (Variable Costs)", {
            x: 0.5, y: 0.3, w: 9, h: 0.5,
            fontSize: 28, color: COLORS.PRIMARY, bold: true, fontFace: FONTS.TITLE
        });

        // Card 1: Labor
        slide.addShape(pres.ShapeType.roundRect, { x: 0.5, y: 1.5, w: 4.2, h: 3.2, fill: "E6F3FF", line: { color: COLORS.SECONDARY, width: 0 } });
        slide.addText("人工成本 (Labor Cost)", { x: 0.7, y: 1.7, w: 3.8, h: 0.4, fontSize: 18, bold: true, color: COLORS.PRIMARY, fontFace: FONTS.BODY });
        
        slide.addText("数据源：HR系统 + 产线IOT打卡", { x: 0.7, y: 2.2, w: 3.8, h: 0.3, fontSize: 14, bold: true, color: COLORS.TEXT_MAIN, fontFace: FONTS.BODY });
        slide.addText([
            { text: "HR系统管理进出厂考勤（获取总工时）。", options: { breakLine: true, bullet: bulletOpts } },
            { text: "产线设备记录具体工位上岗时间（获取直接工时）。", options: { breakLine: true, bullet: bulletOpts } },
            { text: "两者结合实现精确的人员效率分析。", options: { bullet: bulletOpts } }
        ], { x: 0.7, y: 2.6, w: 3.8, h: 2, fontSize: 12, color: COLORS.TEXT_MAIN, fontFace: FONTS.BODY, paraSpaceAfter: 8 });

        // Card 2: Material
        slide.addShape(pres.ShapeType.roundRect, { x: 5.3, y: 1.5, w: 4.2, h: 3.2, fill: "FFF0E6", line: { color: COLORS.ACCENT, width: 0 } });
        slide.addText("物料成本 (Material Cost)", { x: 5.5, y: 1.7, w: 3.8, h: 0.4, fontSize: 18, bold: true, color: COLORS.ACCENT, fontFace: FONTS.BODY });

        slide.addText("数据源：SAP BOM + 扫码枪", { x: 5.5, y: 2.2, w: 3.8, h: 0.3, fontSize: 14, bold: true, color: COLORS.TEXT_MAIN, fontFace: FONTS.BODY });
        slide.addText([
            { text: "基于Manifest中的BOM配置进行校验。", options: { breakLine: true, bullet: bulletOpts } },
            { text: "通过扫码枪验证物料并自动扣减。", options: { breakLine: true, bullet: bulletOpts } },
            { text: "实时计算产线物料消耗与剩余量。", options: { bullet: bulletOpts } }
        ], { x: 5.5, y: 2.6, w: 3.8, h: 2, fontSize: 12, color: COLORS.TEXT_MAIN, fontFace: FONTS.BODY, paraSpaceAfter: 8 });

        // Prompt Note
        slide.addText("Visuals Prompt: Icons representing a worker scanning a badge and a raw material box being scanned.", {
            x: 0.2, y: 5.3, w: 8, h: 0.3, fontSize: 8, color: "BBBBBB"
        });
    }

    // ==========================================================================
    // Slide 6: TOC Data Breakdown - Fixed & Other Costs (Chart)
    // ==========================================================================
    {
        let slide = pres.addSlide({ masterName: "CLEAN_MASTER" });

        slide.addText("TOC数据拆解：固定与隐性成本", {
            x: 0.5, y: 0.3, w: 9, h: 0.5,
            fontSize: 28, color: COLORS.PRIMARY, bold: true, fontFace: FONTS.TITLE
        });

        // Left: Text List
        slide.addText("数据采集挑战与方案", { x: 0.5, y: 1.2, w: 4, h: 0.4, fontSize: 16, bold: true, color: COLORS.PRIMARY, fontFace: FONTS.BODY });
        
        slide.addText([
            { text: "能源 (Energy): 部分新工厂上线智能电表MQTT上传；老旧设备月度分摊。", options: { breakLine: true, bullet: bulletOpts } },
            { text: "运维费用 (Maintenance): 备件消耗多为线下Excel，计划数字化。", options: { breakLine: true, bullet: bulletOpts } },
            { text: "质量成本 (COPQ): 废品/返工目前依赖人工录入。", options: { breakLine: true, bullet: bulletOpts } },
            { text: "固定资产 (Fixed Assets): 按月度固定分摊。", options: { bullet: bulletOpts } }
        ], { x: 0.5, y: 1.7, w: 5, h: 3.5, fontSize: 12, lineSpacing: 20, color: COLORS.TEXT_MAIN, fontFace: FONTS.BODY, paraSpaceAfter: 12 });

        // Right: Chart
        let dataChart = [
            {
                name: "自动化程度",
                labels: ["人工(Labor)", "物料(Mat.)", "能源(Energy)", "运维(Maint.)", "质量(COPQ)"],
                values: [90, 85, 60, 30, 40]
            }
        ];
        
        slide.addChart(pres.ChartType.bar, dataChart, {
            x: 5.8, y: 1.5, w: 4, h: 3.5,
            chartColors: [COLORS.PRIMARY],
            barDir: 'bar',
            title: "各项成本数据实时化程度预估",
            showTitle: true,
            titleFontSize: 14,
            valAxisTitle: "自动化百分比 (%)",
            showValAxisTitle: true,
            catAxisTitle: "成本类别",
            showCatAxisTitle: true,
            valAxisMinVal: 0, valAxisMaxVal: 100
        });

        // Prompt Note
        slide.addText("Visuals Prompt: Digital electric meter displaying live numbers; a warning icon for scrap/maintenance.", {
            x: 0.2, y: 5.3, w: 8, h: 0.3, fontSize: 8, color: "BBBBBB"
        });
    }

    // ==========================================================================
    // Slide 7: System Landscape
    // ==========================================================================
    {
        let slide = pres.addSlide({ masterName: "CLEAN_MASTER" });

        slide.addText("系统集成版图与边界", {
            x: 0.5, y: 0.3, w: 9, h: 0.5,
            fontSize: 28, color: COLORS.PRIMARY, bold: true, fontFace: FONTS.TITLE
        });

        // Top Layer: SAP
        slide.addShape(pres.ShapeType.rect, { x: 2, y: 1.2, w: 6, h: 0.8, fill: "2C3E50" });
        slide.addText("SAP (ERP Core)", { x: 2, y: 1.2, w: 6, h: 0.8, align: "center", color: COLORS.WHITE, bold: true, fontFace: FONTS.BODY });
        slide.addText("管理BOM、计划、财务及主数据", { x: 8.2, y: 1.2, w: 1.6, h: 0.8, fontSize: 10, color: COLORS.TEXT_LIGHT, fontFace: FONTS.BODY });

        // Connector Top-Mid
        slide.addShape(pres.ShapeType.upDownArrow, { x: 4.8, y: 2.1, w: 0.4, h: 0.4, fill: "95A5A6" });

        // Middle Layer: SIM & LESS
        slide.addShape(pres.ShapeType.rect, { x: 2, y: 2.6, w: 2.8, h: 0.8, fill: "34495E" });
        slide.addText("SIM (车间管理)", { x: 2, y: 2.6, w: 2.8, h: 0.8, align: "center", color: COLORS.WHITE, fontFace: FONTS.BODY });
        
        slide.addShape(pres.ShapeType.rect, { x: 5.2, y: 2.6, w: 2.8, h: 0.8, fill: "34495E" });
        slide.addText("LESS (仓储物流)", { x: 5.2, y: 2.6, w: 2.8, h: 0.8, align: "center", color: COLORS.WHITE, fontFace: FONTS.BODY });
        
        slide.addText("SIM: 工单执行 | LESS: 线边物流\n", { x: 8.2, y: 2.6, w: 1.6, h: 0.8, fontSize: 10, color: COLORS.TEXT_LIGHT, fontFace: FONTS.BODY });

        // Connector Mid-Bot
        slide.addShape(pres.ShapeType.upDownArrow, { x: 4.8, y: 3.5, w: 0.4, h: 0.4, fill: "95A5A6" });

        // Bottom Layer: Factory Floor
        slide.addShape(pres.ShapeType.rect, { x: 2, y: 4.0, w: 6, h: 1, fill: COLORS.PRIMARY });
        slide.addText("Factory Floor (MQTT + Edge)", { x: 2, y: 4.0, w: 6, h: 1, align: "center", color: COLORS.WHITE, bold: true, fontFace: FONTS.BODY });
        slide.addText("实时产线数据采集\nKafka/中间表异步交互\n", { x: 8.2, y: 4.0, w: 1.6, h: 1, fontSize: 10, color: COLORS.TEXT_LIGHT, fontFace: FONTS.BODY });

        // Prompt Note
        slide.addText("Visuals Prompt: Integration map. Top: SAP. Middle: SIM & LESS. Bottom: Factory Floor. Connected via bus.", {
            x: 0.2, y: 5.3, w: 8, h: 0.3, fontSize: 8, color: "BBBBBB"
        });
    }

    // ==========================================================================
    // Slide 8: Roadmap
    // ==========================================================================
    {
        let slide = pres.addSlide({ masterName: "CLEAN_MASTER" });

        slide.addText("下一步规划与展望", {
            x: 0.5, y: 0.3, w: 9, h: 0.5,
            fontSize: 28, color: COLORS.PRIMARY, bold: true, fontFace: FONTS.TITLE
        });

        const timelineY = 2.5;
        
        // Step 1
        slide.addShape(pres.ShapeType.chevron, { x: 0.5, y: timelineY, w: 2.8, h: 1.5, fill: COLORS.SECONDARY });
        slide.addText("1. 系统标准化", { x: 0.8, y: timelineY + 0.2, w: 2, h: 0.3, color: COLORS.WHITE, bold: true, fontFace: FONTS.BODY });
        slide.addText("淘汰'简道云'等临时工具，建立统一软件工程规范。", { x: 0.8, y: timelineY + 0.5, w: 2.2, h: 0.8, color: COLORS.WHITE, fontSize: 12, fontFace: FONTS.BODY });

        // Step 2
        slide.addShape(pres.ShapeType.chevron, { x: 3.5, y: timelineY, w: 2.8, h: 1.5, fill: COLORS.PRIMARY });
        slide.addText("2. 全面推广", { x: 3.8, y: timelineY + 0.2, w: 2, h: 0.3, color: COLORS.WHITE, bold: true, fontFace: FONTS.BODY });
        slide.addText("从POC概念验证走向全工厂标准化复制。", { x: 3.8, y: timelineY + 0.5, w: 2.2, h: 0.8, color: COLORS.WHITE, fontSize: 12, fontFace: FONTS.BODY });

        // Step 3
        slide.addShape(pres.ShapeType.chevron, { x: 6.5, y: timelineY, w: 2.8, h: 1.5, fill: COLORS.ACCENT });
        slide.addText("3. AI赋能", { x: 6.8, y: timelineY + 0.2, w: 2, h: 0.3, color: COLORS.WHITE, bold: true, fontFace: FONTS.BODY });
        slide.addText("利用MQTT历史数据训练AI，分析能耗与设备状态相关性。", { x: 6.8, y: timelineY + 0.5, w: 2.2, h: 0.8, color: COLORS.WHITE, fontSize: 12, fontFace: FONTS.BODY });

        // Prompt Note
        slide.addText("Visuals Prompt: Roadmap illustration with stepping stones: 'Standardization', 'Promotion', 'AI Integration'.", {
            x: 0.2, y: 5.3, w: 8, h: 0.3, fontSize: 8, color: "BBBBBB"
        });
    }
}

// 2. Add a Slide to the presentation
buildSlides(pres);

// 3. Save the Presentation
pres.writeFile({ fileName: "Joyson_Digital_Transformation.pptx" });