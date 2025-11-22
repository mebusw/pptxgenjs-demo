// pnpm install pptxgenjs
// nvm use 22
import pptxgen from "pptxgenjs";

// 1. Create a Presentation
let pres = new pptxgen();

pres.title = 'My Awesome Presentation';
pres.author = 'Brent Ely';
pres.subject = 'Annual Report';
pres.company = 'Computer Science Chair';
pres.revision = '15';
//Default Font 
pres.theme = { headFontFace: "Arial Light" };
pres.theme = { bodyFontFace: "Arial" };
// Define new layout for the Presentation
pres.defineLayout({ name: 'A3', width: 16.5, height: 11.7 });
// Set presentation to use new layout
pres.layout = 'A3';

pres.defineSlideMaster({
    title: "MASTER_SLIDE",
    background: { color: "cccccc" },
    objects: [
        { line: { x: 3.5, y: 1.0, w: 6.0, line: { color: "0088CC", width: 5 } } },
        { rect: { x: 0.0, y: 5.3, w: "100%", h: 0.75, fill: { color: "F1F1F1" } } },
        { text: { text: "Status Report", options: { x: 3.0, y: 5.3, w: 5.5, h: 0.75 } } },
        { image: { x: 12, y: 1, w: 3.2, h: 0.75, path: "images/logo.png" } },
        {
            placeholder: {
                options: { name: "title", type: "title", x: 0.6, y: 0.5, w: 12, h: 1.25 },
                text: "(custom placeholder title!)",
            },
        },
        {
            placeholder: {
                options: { name: "body", type: "body", x: 0.6, y: 1.5, w: 12, h: 5.25 },
                text: "(custom placeholder body!)",
            },
        },
    ],
    slideNumber: { x: 0.3, y: "90%" },
});


// 2. Add a Slide to the presentation
let slide1 = pres.addSlide({ masterName: "MASTER_SLIDE" });
slide1.addText("Slide 1 大标题", { placeholder: "title" });
slide1.addText("Body Placeholder here!", { placeholder: "body" });
slide1.background = { color: "FF3399", transparency: 50 }; // hex fill color with transparency of 50%
// slide.background = { data: "image/png;base64,ABC[...]123" }; // image: base64 data
// slide.background = { path: "https://some.url/image.jpg" }; // image: url
slide1.color = "696969";
slide1.addText("Hello World", { x: 1, y: 1, fontSize: 24, color: "363636" });
slide1.addImage({ path: "cats-fight.jpeg", x: 2, y: 5, w: 6, h: 4 });
// slide1.addMedia({ type: "video", path: "https://www.youtube.com/embed/Dph6ynRVyUc", x: 10, y: 8, w: 5, h: 3 });

// 2. Add objects (Tables, Shapes, etc.) to the Slide
let slide2 = pres.addSlide({ masterName: "MASTER_SLIDE" });
slide2.addText("Slide Two", { x: 1, y: 1 });
let textboxText = "Hello World from PptxGenJS!";
let textboxOpts = { x: 1, y: 3, color: "363636" };
slide2.addText(textboxText, textboxOpts);
// Shapes without text
slide2.addShape(pres.ShapeType.rect, { fill: { color: "FF0000" }, x: 5, y: 4, });
slide2.addShape(pres.ShapeType.ellipse, {
    fill: { type: "solid", color: "0088CC" }, x: 5, y: 5,
});
slide2.addShape(pres.ShapeType.line, { line: { color: "FF0000", width: 1 }, x: 5, y: 6, });
// Shapes with text
slide2.addText("ShapeType.rect", {
    shape: pres.ShapeType.rect,
    fill: { color: "0088CC" },
    x: 5, y: 4,
});
slide2.addText("ShapeType.ellipse", {
    shape: pres.ShapeType.ellipse,
    fill: { color: "0088CC" },
    x: 5, y: 5,
});
slide2.addText("ShapeType.line", {
    shape: pres.ShapeType.line,
    line: { color: "0088CC", width: 1, dashType: "lgDash" },
    x: 5, y: 6,
});


// 2. Add Charts to presentation
let slide3 = pres.addSlide({ masterName: "MASTER_SLIDE" });
let dataChartAreaLine = [
    {
        name: "Actual Sales",
        labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
        values: [1500, 4600, 5156, 3167, 8510, 8009, 6006, 7855, 12102, 12789, 10123, 15121],
    },
    {
        name: "Projected Sales",
        labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
        values: [1000, 2600, 3456, 4567, 5010, 6009, 7006, 8855, 9102, 10789, 11123, 12121],
    },
];
slide3.addChart(pres.ChartType.line, dataChartAreaLine, { x: 1, y: 1, w: 8, h: 4 });
let rows = [
    [
        { text: "Top Lft", options: { align: "left", fontFace: "Arial" } },
        { text: "Top Ctr", options: { align: "center", fontFace: "Verdana" } },
        { text: "Top Rgt", options: { align: "right", fontFace: "Courier" } },
    ],
    [
        { text: "Bottom Lft", options: { align: "left", fontFace: "Arial" } },
        { text: "Bottom Ctr", options: { align: "center", fontFace: "Verdana" } },
        { text: "Bottom Rgt", options: { align: "right", fontFace: "Courier" } },
    ],
];
slide3.addTable(rows, { w: 9, rowH: 1, align: "left", fontFace: "Arial", x: 1, y: 6 });

// 3. Save the Presentation
pres.writeFile({ fileName: "Hello-World.pptx" });