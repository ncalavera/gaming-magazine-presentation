const pptxgen = require('pptxgenjs');
const html2pptx = require('./html2pptx');
const path = require('path');

async function createPresentation() {
    try {
        // Create new presentation
        const pptx = new pptxgen();
        pptx.layout = 'LAYOUT_16x9';  // Must match HTML body dimensions (720pt × 405pt)
        pptx.author = 'Futura Team';
        pptx.title = 'Журнал для игровой индустрии';

        console.log('Creating presentation...');

        // Process each slide
        const slides = [
            'slides/slide1.html',  // Title slide
            'slides/slide2.html',  // Market context
            'slides/slide3.html',  // Magazine concept
            'slides/slide4.html',  // Issue structure
            'slides/slide5.html',  // Editorial team
            'slides/slide6.html',  // Project economics
            'slides/slide7.html'   // Partnership conclusion
        ];

        for (let i = 0; i < slides.length; i++) {
            const slideFile = path.join(__dirname, slides[i]);
            console.log(`Processing slide ${i + 1}: ${slides[i]}`);

            const { slide, placeholders } = await html2pptx(slideFile, pptx);

            // No charts or tables to add for these slides - they're all content slides
            console.log(`Slide ${i + 1} created successfully`);
        }

        // Save the presentation
        const outputPath = path.join(__dirname, 'Журнал_для_игровой_индустрии.pptx');
        await pptx.writeFile({ fileName: outputPath });

        console.log(`Presentation saved to: ${outputPath}`);
        console.log('✅ Presentation created successfully!');

    } catch (error) {
        console.error('❌ Error creating presentation:', error);
        throw error;
    }
}

// Run the presentation creation
createPresentation().catch(console.error);