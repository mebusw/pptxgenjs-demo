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
    ],
    slideNumber: { x: 0.3, y: "90%" },
});

let buildSlides = (pres) => {
    let slide1 = pres.addSlide({ masterName: "MASTER_SLIDE" });
    console.log("building slide1...");
    //...

}

// 2. Add a Slide to the presentation
buildSlides(pres);

// 3. Save the Presentation
pres.writeFile({ fileName: "Hello-World.pptx" });