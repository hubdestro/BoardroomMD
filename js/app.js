/**
 * App Entry Point
 * Handles UI interactions, events, and strictly decouples DOM from logic.
 */

document.addEventListener('DOMContentLoaded', () => {
    // UI Elements
    const generateBtn = document.getElementById('generateBtn');
    const markdownInput = document.getElementById('markdownInput');
    const docTitleInput = document.getElementById('docTitle');
    const docAuthorInput = document.getElementById('docAuthor');
    const statusMessage = document.getElementById('statusMessage');
    const templateRadios = document.querySelectorAll('input[name="template"]');
    
    // Tab Elements
    const tabEditor = document.getElementById('tabEditor');
    const tabPreview = document.getElementById('tabPreview');
    const previewContainer = document.getElementById('previewContainer');
    const docxPreviewEl = document.getElementById('docxPreview');
    const previewLoader = document.getElementById('previewLoader');

    // UI Enhancements: Template selection styling
    templateRadios.forEach(radio => {
        radio.addEventListener('change', (e) => {
            // Remove selected class from all
            document.querySelectorAll('.template-radio-container').forEach(container => {
                container.classList.remove('border-brand-500', 'bg-brand-50');
                container.classList.add('border-gray-200');
            });
            
            // Add to currently selected
            if (e.target.checked) {
                const container = e.target.closest('.template-radio-container');
                container.classList.remove('border-gray-200');
                container.classList.add('border-brand-500', 'bg-brand-50');
            }
        });
    });

    // Tab Switching Logic
    function switchTab(tab) {
        // Clear any leftover inline styles from previous experiments
        markdownInput.style.height = '';
        markdownInput.style.flex = '';
        previewContainer.style.height = '';
        previewContainer.style.flex = '';

        if (tab === 'editor') {
            tabEditor.classList.add('text-brand-600', 'border-brand-600', 'font-bold');
            tabEditor.classList.remove('text-gray-400', 'border-transparent', 'font-semibold', 'hover:border-brand-300');
            tabPreview.classList.add('text-gray-400', 'border-transparent', 'font-semibold', 'hover:border-brand-300');
            tabPreview.classList.remove('text-brand-600', 'border-brand-600', 'font-bold');
            
            markdownInput.classList.remove('hidden');
            markdownInput.classList.add('flex-1');
            previewContainer.classList.add('hidden');
            previewContainer.classList.remove('flex-1');
        } else {
            tabPreview.classList.add('text-brand-600', 'border-brand-600', 'font-bold');
            tabPreview.classList.remove('text-gray-400', 'border-transparent', 'font-semibold', 'hover:border-brand-300');
            tabEditor.classList.add('text-gray-400', 'border-transparent', 'font-semibold', 'hover:border-brand-300');
            tabEditor.classList.remove('text-brand-600', 'border-brand-600', 'font-bold');
            
            markdownInput.classList.add('hidden');
            markdownInput.classList.remove('flex-1');
            
            previewContainer.classList.remove('hidden');
            previewContainer.classList.add('flex-1');
            renderPreview();
        }
    }

    tabEditor.addEventListener('click', () => switchTab('editor'));
    tabPreview.addEventListener('click', () => switchTab('preview'));


    // Real-time preview update
    let debounceTimer;
    markdownInput.addEventListener('input', () => {
        if (!previewContainer.classList.contains('hidden')) {
            clearTimeout(debounceTimer);
            debounceTimer = setTimeout(renderPreview, 1000);
        }
    });

    // Re-render preview if template changes while in preview tab
    templateRadios.forEach(radio => {
        radio.addEventListener('change', () => {
            if (!previewContainer.classList.contains('hidden')) {
                renderPreview();
            }
        });
    });

    // Helper: Core document generation flow
    async function getDocumentBlob() {
        const rawMarkdown = markdownInput.value.trim();
        if (!rawMarkdown) {
            throw new Error("Please enter some Markdown text.");
        }

        // Try to extract title from Markdown if empty
        let docTitle = docTitleInput.value.trim();
        if (!docTitle) {
            const match = rawMarkdown.match(/^#\s+(.+)$/m);
            docTitle = match ? match[1] : "Generated Document";
        }
        
        const docAuthor = docAuthorInput.value.trim() || "Author";
        const selectedTemplate = document.querySelector('input[name="template"]:checked').value;

        // Parse & Map
        const tokens = window.MarkdownParser.parse(rawMarkdown);
        const docxElements = window.TokenMapper.mapTokensToDocx(tokens);

        // Generate Blob
        return await window.DocumentGenerator.getBlob(
            docxElements, 
            selectedTemplate, 
            { title: docTitle, author: docAuthor }
        );
    }

    // Preview Action
    async function renderPreview() {
        previewLoader.classList.remove('hidden');
        docxPreviewEl.innerHTML = ''; // Clear previous preview
        // previewContainer.scrollTop = 0; // Don't reset scroll for live updates
        
        try {
            const blob = await getDocumentBlob();
            
            // Render using docx-preview
            await window.docxPreview.renderAsync(blob, docxPreviewEl);
            
        } catch (error) {
            console.error("Preview failed:", error);
            docxPreviewEl.innerHTML = `<div class="p-8 text-center"><div class="text-red-500 font-medium">Error rendering preview: ${error.message}</div></div>`;
        } finally {
            previewLoader.classList.add('hidden');
        }
    }

    // Generate & Download Action
    generateBtn.addEventListener('click', async () => {
        generateBtn.disabled = true;
        generateBtn.classList.add('opacity-75', 'cursor-not-allowed');
        statusMessage.classList.remove('hidden');
        statusMessage.textContent = "Generating DOCX File...";

        try {
            const blob = await getDocumentBlob();
            const docTitle = docTitleInput.value.trim() || "Generated Document";
            
            window.DocumentGenerator.download(blob, { title: docTitle });

            statusMessage.textContent = "Download Complete!";
            setTimeout(() => statusMessage.classList.add('hidden'), 3000);

        } catch (error) {
            console.error("Workflow failed:", error);
            statusMessage.textContent = error.message;
            statusMessage.classList.add('text-red-500');
            setTimeout(() => {
                statusMessage.classList.remove('text-red-500');
                statusMessage.classList.add('hidden');
            }, 3000);
        } finally {
            generateBtn.disabled = false;
            generateBtn.classList.remove('opacity-75', 'cursor-not-allowed');
        }
    });
});
