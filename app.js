// Office 初始化
Office.onReady(function(info) {
    console.log("Office.js 已加载，插件初始化完成。");
});

/**
 * 主转换函数，由 Ribbon 按钮调用
 */
async function convertMarkdownToWord() {
    try {
        await Word.run(async (context) => {
            // 获取选中的 Markdown 文本
            let selectedTextRange = context.document.getSelection();
            selectedTextRange.load('text');
            await context.sync();

            let markdownText = selectedTextRange.text;

            if (!markdownText.trim()) {
                console.log("未选中任何文本或选中文本为空。");
                displayMessage("请选中 Markdown 文本后重试。");
                return;
            }

            // 解析 Markdown 文本
            const parsedElements = parseMarkdownText(markdownText);

            // 删除选中文本并保持选中状态
            selectedTextRange.delete('Select');
            await context.sync();

            let currentRange = context.document.getSelection();

            // 根据解析结果应用 Word 格式
            await applyWordFormatting(context, currentRange, parsedElements);

            await context.sync();
            console.log("Markdown 转换完成！");
            displayMessage("Markdown 转换成功完成！");
        });
    } catch (error) {
        console.error("转换过程中发生错误:", error);
        displayMessage("转换失败: " + error.message);
    }
}

/**
 * 解析 Markdown 文本，使用 marked.js 库或自定义逻辑
 * @param {string} markdownText Markdown 文本
 * @returns {Array<object>} 解析后的元素数组
 */
function parseMarkdownText(markdownText) {
    const elements = [];
    const lines = markdownText.split('\n');

    // 简化解析逻辑，实际项目中建议使用 marked.js 解析
    lines.forEach(line => {
        line = line.trim();
        if (!line) {
            elements.push({ type: 'paragraph', text: '' });
            return;
        }

        if (line.startsWith('# ')) {
            elements.push({ type: 'heading1', text: line.substring(2) });
        } else if (line.startsWith('## ')) {
            elements.push({ type: 'heading2', text: line.substring(3) });
        } else if (line.startsWith('### ')) {
            elements.push({ type: 'heading3', text: line.substring(4) });
        } else if (line.startsWith('- ') || line.startsWith('* ')) {
            elements.push({ type: 'listItem', text: line.substring(2) });
        } else if (line.match(/\*\*.*?\*\*/)) {
            let text = line;
            const boldText = text.match(/\*\*(.*?)\*\*/);
            elements.push({
                type: 'paragraph',
                text: text.replace(/\*\*(.*?)\*\*/g, '$1'),
                format: { bold: boldText ? boldText[1] : '' }
            });
        } else if (line.match(/\*.*?\*/)) {
            let text = line;
            const italicText = text.match(/\*(.*?)\*/);
            elements.push({
                type: 'paragraph',
                text: text.replace(/\*(.*?)\*/g, '$1'),
                format: { italic: italicText ? italicText[1] : '' }
            });
        } else {
            elements.push({ type: 'paragraph', text: line });
        }
    });

    return elements;
}

/**
 * 应用 Word 格式化
 * @param {Word.RequestContext} context Word 请求上下文
 * @param {Word.Range} range 要插入内容的 Word 范围
 * @param {Array<object>} parsedElements 解析后的 Markdown 元素数组
 */
async function applyWordFormatting(context, range, parsedElements) {
    for (const element of parsedElements) {
        switch (element.type) {
            case 'heading1':
                let heading1Paragraph = range.insertParagraph(element.text, Word.InsertLocation.end);
                heading1Paragraph.style = "Heading 1";
                range = heading1Paragraph;
                break;
            case 'heading2':
                let heading2Paragraph = range.insertParagraph(element.text, Word.InsertLocation.end);
                heading2Paragraph.style = "Heading 2";
                range = heading2Paragraph;
                break;
            case 'heading3':
                let heading3Paragraph = range.insertParagraph(element.text, Word.InsertLocation.end);
                heading3Paragraph.style = "Heading 3";
                range = heading3Paragraph;
                break;
            case 'paragraph':
                let paragraph = range.insertParagraph(element.text, Word.InsertLocation.end);
                if (element.format) {
                    if (element.format.bold) {
                        let boldRange = paragraph.getTextRanges([element.format.bold], true);
                        boldRange.font.bold = true;
                    }
                    if (element.format.italic) {
                        let italicRange = paragraph.getTextRanges([element.format.italic], true);
                        italicRange.font.italic = true;
                    }
                }
                range = paragraph;
                break;
            case 'listItem':
                let listItemParagraph = range.insertParagraph(element.text, Word.InsertLocation.end);
                listItemParagraph.listItem.level = 0;
                listItemParagraph.listItem.applyBullet();
                range = listItemParagraph;
                break;
            default:
                range = range.insertParagraph(element.text || '', Word.InsertLocation.end);
                break;
        }
    }
}

/**
 * 显示消息（可选，用于调试或用户提示）
 * @param {string} message 消息内容
 */
function displayMessage(message) {
    const messageElement = document.getElementById("message");
    if (messageElement) {
        messageElement.textContent = message;
        messageElement.style.display = "block";
    }
}
