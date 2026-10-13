'use client';

import React, { useState, useRef, useEffect, useMemo, useCallback } from 'react';
import { useLanguage } from './LanguageContext';
import { markdownToHtml, SAMPLE_MARKDOWN_PRESETS } from '@/lib/markdown-to-html';
import { renderLatexHtml } from '@/lib/latex-to-omml';
import { importDocxFile, htmlToMarkdown } from '@/lib/docx-import';
import { WordRibbon, WordFormattingOptions, ViewMode } from './WordRibbon';
import { WordRuler } from './WordRuler';
import { DocumentStatsModal } from './DocumentStatsModal';
import { FindReplaceModal } from './FindReplaceModal';
import { InsertImageModal } from './InsertImageModal';
import { InsertLinkModal } from './InsertLinkModal';
import { SymbolPaletteModal } from './SymbolPaletteModal';
import { WatermarkModal } from './WatermarkModal';
import {
  FileText,
  Eye,
  FileCode2,
  Split,
  Table as TableIcon,
  Sigma,
  Code,
  Upload,
  Check,
  Sparkles,
  ArrowRight,
  Layers,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';

interface EditorProps {
  markdown: string;
  onChangeMarkdown: (markdown: string) => void;
  documentTitle: string;
  onChangeTitle: (title: string) => void;
  onOpenInsertModal: () => void;
  onExportDocx?: () => void;
  onExportPdf?: () => void;
  onCopyGoogleDocs?: () => void;
  onSaveDoc?: () => void;
  onClearDoc?: () => void;
  isExporting?: boolean;
}

export const Editor: React.FC<EditorProps> = ({
  markdown,
  onChangeMarkdown,
  documentTitle,
  onChangeTitle,
  onOpenInsertModal,
  onExportDocx = () => {},
  onExportPdf = () => {},
  onCopyGoogleDocs = () => {},
  onSaveDoc = () => {},
  onClearDoc = () => {},
  isExporting = false,
}) => {
  const { direction, t } = useLanguage();

  // Word Editor Formatting & Layout Settings
  const [formatting, setFormatting] = useState<WordFormattingOptions>({
    fontFamily: 'Segoe UI',
    fontSize: 11,
    orientation: 'portrait',
    paperSize: 'a4',
    margins: 'normal',
    lineSpacing: '1.15',
    watermark: '',
    showRuler: true,
    zoomLevel: 100,
  });

  const [viewMode, setViewMode] = useState<ViewMode>('print');
  const [history, setHistory] = useState<string[]>([markdown]);
  const [historyIndex, setHistoryIndex] = useState(0);

  // Drag & drop status for .docx files
  const [isDraggingDocx, setIsDraggingDocx] = useState(false);
  const [importNotification, setImportNotification] = useState<string | null>(null);

  // Modals state
  const [showStatsModal, setShowStatsModal] = useState(false);
  const [showFindReplaceModal, setShowFindReplaceModal] = useState(false);
  const [showImageModal, setShowImageModal] = useState(false);
  const [showLinkModal, setShowLinkModal] = useState(false);
  const [showSymbolModal, setShowSymbolModal] = useState(false);
  const [showWatermarkModal, setShowWatermarkModal] = useState(false);

  // Math, Table, Code dialogs
  const [showMathDialog, setShowMathDialog] = useState(false);
  const [mathInput, setMathInput] = useState('\\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}');
  const [showTableDialog, setShowTableDialog] = useState(false);
  const [tableRowsCount, setTableRowsCount] = useState(3);
  const [tableColsCount, setTableColsCount] = useState(3);
  const [showCodeDialog, setShowCodeDialog] = useState(false);
  const [codeLang, setCodeLang] = useState('python');
  const [codeSnippet, setCodeSnippet] = useState('def calculate_result():\n    return 42');

  // References
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const splitTextareaRef = useRef<HTMLTextAreaElement>(null);
  const paperEditableRef = useRef<HTMLDivElement>(null);
  const savedSelectionRange = useRef<Range | null>(null);
  const isTypingDirectlyRef = useRef(false);
  const isSyncingRef = useRef(false);

  // Save selection range on mouseup/keyup in browser
  const saveSelection = useCallback(() => {
    if (typeof window !== 'undefined') {
      const sel = window.getSelection();
      if (sel && sel.rangeCount > 0) {
        savedSelectionRange.current = sel.getRangeAt(0).cloneRange();
      }
    }
  }, []);

  // Restore selection range
  const restoreSelection = useCallback(() => {
    if (typeof window !== 'undefined' && savedSelectionRange.current) {
      const sel = window.getSelection();
      if (sel) {
        try {
          sel.removeAllRanges();
          sel.addRange(savedSelectionRange.current);
        } catch {}
      }
    }
  }, []);

  // Update markdown with history
  const updateMarkdownWithHistory = useCallback((newMd: string) => {
    onChangeMarkdown(newMd);
    setHistory(prev => {
      const next = prev.slice(0, historyIndex + 1);
      next.push(newMd);
      return next;
    });
    setHistoryIndex(prev => prev + 1);
  }, [historyIndex, onChangeMarkdown]);

  // Undo / Redo handlers
  const handleUndo = useCallback(() => {
    if (historyIndex > 0) {
      const newIdx = historyIndex - 1;
      setHistoryIndex(newIdx);
      const prevMd = history[newIdx];
      onChangeMarkdown(prevMd);
      if (paperEditableRef.current) {
        paperEditableRef.current.innerHTML = markdownToHtml(prevMd);
      }
    }
  }, [history, historyIndex, onChangeMarkdown]);

  const handleRedo = useCallback(() => {
    if (historyIndex < history.length - 1) {
      const newIdx = historyIndex + 1;
      setHistoryIndex(newIdx);
      const nextMd = history[newIdx];
      onChangeMarkdown(nextMd);
      if (paperEditableRef.current) {
        paperEditableRef.current.innerHTML = markdownToHtml(nextMd);
      }
    }
  }, [history, historyIndex, onChangeMarkdown]);

  // Keep paper content in sync with markdown when not actively typing or syncing
  useEffect(() => {
    if (!isSyncingRef.current && !isTypingDirectlyRef.current && paperEditableRef.current) {
      const rendered = markdownToHtml(markdown);
      if (paperEditableRef.current.innerHTML !== rendered) {
        paperEditableRef.current.innerHTML = rendered;
      }
    }
    isSyncingRef.current = false;
  }, [markdown]);

  // Sync paper DOM back into markdown
  const syncPaperToMarkdown = useCallback(() => {
    if (paperEditableRef.current) {
      const html = paperEditableRef.current.innerHTML;
      const md = htmlToMarkdown(html);
      if (md && md.trim() !== markdown.trim()) {
        isSyncingRef.current = true;
        updateMarkdownWithHistory(md);
      }
    }
  }, [markdown, updateMarkdownWithHistory]);

  const handlePaperInput = (e: React.FormEvent<HTMLDivElement>) => {
    isTypingDirectlyRef.current = true;
    saveSelection();
    const html = e.currentTarget.innerHTML;
    const md = htmlToMarkdown(html);
    if (md) {
      onChangeMarkdown(md);
    }
  };

  const handlePaperBlur = (e: React.FocusEvent<HTMLDivElement>) => {
    isTypingDirectlyRef.current = false;
    syncPaperToMarkdown();
  };

  // Central Tool Command Execution Handler (Word Ribbon Tools)
  const handleExecCommand = useCallback((
    command: string,
    value?: string,
    mdBefore?: string,
    mdAfter?: string
  ) => {
    // 1. Textarea in Source mode or Split mode
    const activeTextarea =
      viewMode === 'source'
        ? textareaRef.current
        : viewMode === 'split'
        ? splitTextareaRef.current
        : null;

    if (
      activeTextarea &&
      (document.activeElement === activeTextarea ||
        activeTextarea.selectionStart !== activeTextarea.selectionEnd)
    ) {
      const start = activeTextarea.selectionStart;
      const end = activeTextarea.selectionEnd;
      const selected = markdown.substring(start, end);
      const before =
        mdBefore !== undefined
          ? mdBefore
          : command === 'bold'
          ? '**'
          : command === 'italic'
          ? '*'
          : command === 'underline'
          ? '<u>'
          : '';
      const after =
        mdAfter !== undefined
          ? mdAfter
          : command === 'underline'
          ? '</u>'
          : before;
      const replacement = `${before}${selected || 'Formatted Text'}${after}`;
      const updated = markdown.substring(0, start) + replacement + markdown.substring(end);
      updateMarkdownWithHistory(updated);
      setTimeout(() => {
        activeTextarea.focus();
        activeTextarea.setSelectionRange(start + before.length, end + before.length);
      }, 50);
      return;
    }

    // 2. Interactive Paper Canvas with Selection / Cursor
    if (paperEditableRef.current) {
      paperEditableRef.current.focus();
    }
    restoreSelection();

    let handledInPaper = false;

    if (command && typeof document !== 'undefined') {
      try {
        // Execute document command in the contentEditable paper
        handledInPaper = document.execCommand(command, false, value);
        if (handledInPaper && paperEditableRef.current) {
          syncPaperToMarkdown();
          saveSelection();
          return;
        }
      } catch {}
    }

    // 3. Fallback: Check window selection and wrap matching markdown text or insert HTML
    if (typeof window !== 'undefined') {
      const sel = window.getSelection();
      const selectedText = sel ? sel.toString() : '';

      if (selectedText && selectedText.trim().length > 0) {
        const trimmed = selectedText.trim();
        const before =
          mdBefore !== undefined
            ? mdBefore
            : command === 'bold'
            ? '**'
            : command === 'italic'
            ? '*'
            : command === 'underline'
            ? '<u>'
            : '';
        const after =
          mdAfter !== undefined
            ? mdAfter
            : command === 'underline'
            ? '</u>'
            : before;

        if (markdown.includes(trimmed)) {
          const updated = markdown.replace(trimmed, `${before}${trimmed}${after}`);
          updateMarkdownWithHistory(updated);
          return;
        } else {
          // Wrap in DOM directly
          try {
            document.execCommand(
              'insertHTML',
              false,
              `${before}${trimmed}${after}`
            );
            syncPaperToMarkdown();
            return;
          } catch {}
        }
      }
    }

    // 4. Default: Wrap or append formatted syntax
    const before =
      mdBefore !== undefined
        ? mdBefore
        : command === 'bold'
        ? '**'
        : command === 'italic'
        ? '*'
        : command === 'underline'
        ? '<u>'
        : '';
    const after =
      mdAfter !== undefined
        ? mdAfter
        : command === 'underline'
        ? '</u>'
        : before;

    // Try inserting HTML into paper at cursor
    if (paperEditableRef.current) {
      try {
        const placeholder =
          command === 'bold'
            ? '<b>Bold Text</b>'
            : command === 'italic'
            ? '<i>Italic Text</i>'
            : command === 'underline'
            ? '<u>Underlined Text</u>'
            : `${before}Formatted Text${after}`;
        document.execCommand('insertHTML', false, placeholder);
        syncPaperToMarkdown();
        return;
      } catch {}
    }

    updateMarkdownWithHistory(`${markdown}\n\n${before}Formatted Text${after}`);
  }, [
    markdown,
    restoreSelection,
    saveSelection,
    syncPaperToMarkdown,
    updateMarkdownWithHistory,
    viewMode,
  ]);

  const updateFormatting = (opts: Partial<WordFormattingOptions>) => {
    setFormatting(prev => ({ ...prev, ...opts }));
  };

  // Helper to wrap selected text in active textarea or paper
  const wrapSelection = (before: string, after: string = before) => {
    handleExecCommand('', undefined, before, after);
  };

  // Apply block-level prefix (headings, lists, blockquotes)
  const applyBlockStyle = (prefix: string) => {
    updateMarkdownWithHistory(`${markdown}\n\n${prefix} `);
  };

  // Insert Table
  const handleInsertTable = () => {
    let mdTable = '\n\n|';
    for (let c = 1; c <= tableColsCount; c++) mdTable += ` Column ${c} |`;
    mdTable += '\n|';
    for (let c = 1; c <= tableColsCount; c++) mdTable += ' :--- |';
    for (let r = 1; r <= tableRowsCount; r++) {
      mdTable += '\n|';
      for (let c = 1; c <= tableColsCount; c++) mdTable += ` Data ${r}-${c} |`;
    }
    mdTable += '\n\n';
    updateMarkdownWithHistory(markdown + mdTable);
    setShowTableDialog(false);
  };

  // Insert Math Formula
  const handleInsertMath = () => {
    if (mathInput.trim()) {
      const mdMath = `\n\n$$\n${mathInput.trim()}\n$$\n\n`;
      updateMarkdownWithHistory(markdown + mdMath);
      setShowMathDialog(false);
    }
  };

  // Insert Code Block
  const handleInsertCode = () => {
    const mdCode = `\n\n\`\`\`${codeLang}\n${codeSnippet}\n\`\`\`\n\n`;
    updateMarkdownWithHistory(markdown + mdCode);
    setShowCodeDialog(false);
  };

  // DOCX Import File Handler
  const handleImportDocx = async (file: File) => {
    try {
      setImportNotification('Reading and converting Word document (.docx)...');
      const result = await importDocxFile(file);
      updateMarkdownWithHistory(result.markdown);
      onChangeTitle(result.title);
      setImportNotification(`Successfully imported: "${file.name}" with rich formatting!`);
      setTimeout(() => setImportNotification(null), 4000);
    } catch (err) {
      console.error('Failed to import DOCX:', err);
      setImportNotification('Failed to parse .docx file. Please check file format.');
      setTimeout(() => setImportNotification(null), 4000);
    }
  };

  // Drag and drop events for .docx files
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDraggingDocx(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDraggingDocx(false);
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    setIsDraggingDocx(false);
    const files = Array.from(e.dataTransfer.files);
    const docxFile = files.find(f => f.name.endsWith('.docx'));
    if (docxFile) {
      await handleImportDocx(docxFile);
    }
  };

  // Load sample template from Ribbon
  const handleLoadTemplate = (key: string) => {
    const preset =
      SAMPLE_MARKDOWN_PRESETS.find(p => p.id === key) ||
      (key === 'saudi_vat' ? SAMPLE_MARKDOWN_PRESETS[1] : SAMPLE_MARKDOWN_PRESETS[0]);
    if (preset) {
      updateMarkdownWithHistory(preset.markdown);
      onChangeTitle(preset.titleEn);
    }
  };

  // Calculate statistics
  const cleanText = markdown.replace(/[#*`~>|\[\]\-_$]/g, ' ');
  const words = cleanText.trim().split(/\s+/).filter(Boolean).length;
  const characters = markdown.length;
  const readingMinutes = Math.max(1, Math.ceil(words / 200));

  // Multi-page splitting count
  const pageBreaksCount = (markdown.match(/<div style="page-break-after: always[\s\S]*?<\/div>|\n---\n/gi) || []).length;
  const totalPages = Math.max(1, pageBreaksCount + 1);

  // Compute CSS styles based on Word Formatting Options
  const paperMarginClasses =
    formatting.margins === 'narrow'
      ? 'p-3.5 sm:p-6 md:p-8'
      : formatting.margins === 'wide'
      ? 'p-5 sm:p-10 md:p-16'
      : 'p-4 sm:p-8 md:p-12';

  const paperOrientationWidth =
    formatting.orientation === 'landscape'
      ? 'w-full max-w-[1100px] min-h-[400px] sm:min-h-[750px]'
      : 'w-full max-w-[850px] min-h-[450px] sm:min-h-[1050px]';

  const fontFamilyStyle: React.CSSProperties = {
    fontFamily: formatting.fontFamily,
    fontSize: `${formatting.fontSize}pt`,
    lineHeight: formatting.lineSpacing,
    transform: `scale(${formatting.zoomLevel / 100})`,
    transformOrigin: 'top center',
  };

  return (
    <div
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      className="relative w-full flex flex-col items-center py-2 px-2 sm:px-6"
    >
      {/* Drag & Drop Overlay */}
      {isDraggingDocx && (
        <div className="fixed inset-0 z-50 bg-blue-600/80 backdrop-blur-sm flex flex-col items-center justify-center text-white pointer-events-none animate-fadeIn">
          <Upload className="h-16 w-16 mb-4 animate-bounce" />
          <h2 className="text-2xl font-black mb-2">Drop Word Document (.docx) Here</h2>
          <p className="text-sm text-blue-100 max-w-md text-center">
            MassiveMark will immediately import headings, tables, LaTeX math, and rich formatting!
          </p>
        </div>
      )}

      {/* Import Notification Banner */}
      {importNotification && (
        <div className="no-print w-full max-w-5xl mb-3 p-3 rounded-xl bg-blue-50 dark:bg-blue-950/60 border border-blue-200 dark:border-blue-900 text-blue-800 dark:text-blue-200 text-xs sm:text-sm font-semibold flex items-center justify-between shadow-xs animate-slideDown">
          <div className="flex items-center gap-2">
            <Check className="h-4 w-4 text-blue-600 dark:text-blue-400" />
            <span>{importNotification}</span>
          </div>
          <button
            onClick={() => setImportNotification(null)}
            className="text-xs hover:underline text-blue-600 cursor-pointer"
          >
            Dismiss
          </button>
        </div>
      )}

      {/* 3-Step Workflow Banner */}
      <div className="no-print w-full max-w-5xl mb-3 bg-gradient-to-r from-red-50 via-rose-50 to-orange-50 dark:from-red-950/20 dark:via-rose-950/20 dark:to-orange-950/20 border border-red-200/80 dark:border-red-900/40 rounded-2xl p-3 sm:p-4 shadow-xs">
        <div className="flex flex-col sm:flex-row items-center justify-between gap-3 text-xs sm:text-sm">
          {/* Step 1 */}
          <div className="flex items-center space-x-2 rtl:space-x-reverse text-slate-800 dark:text-slate-200">
            <span className="h-6 w-6 rounded-full bg-red-600 text-white font-bold text-xs flex items-center justify-center shrink-0">
              1
            </span>
            <div>
              <span className="font-bold">{t.step1Title}</span>
              <p className="text-[11px] text-slate-500 dark:text-slate-400 hidden md:block">
                {t.step1Desc}
              </p>
            </div>
          </div>

          <ArrowRight className="h-4 w-4 text-red-400 hidden sm:block rtl:rotate-180" />

          {/* Step 2 */}
          <div
            onClick={onOpenInsertModal}
            className="flex items-center space-x-2 rtl:space-x-reverse text-slate-800 dark:text-slate-200 cursor-pointer group"
          >
            <span className="h-6 w-6 rounded-full bg-red-600 text-white font-bold text-xs flex items-center justify-center shrink-0 group-hover:scale-110 transition-transform">
              2
            </span>
            <div>
              <span className="font-bold text-red-600 dark:text-red-400 group-hover:underline flex items-center gap-1">
                {t.step2Title}
                <Sparkles className="h-3.5 w-3.5" />
              </span>
              <p className="text-[11px] text-slate-500 dark:text-slate-400 hidden md:block">
                {t.step2Desc}
              </p>
            </div>
          </div>

          <ArrowRight className="h-4 w-4 text-red-400 hidden sm:block rtl:rotate-180" />

          {/* Step 3 */}
          <div className="flex items-center space-x-2 rtl:space-x-reverse text-slate-800 dark:text-slate-200">
            <span className="h-6 w-6 rounded-full bg-slate-900 dark:bg-white text-white dark:text-slate-900 font-bold text-xs flex items-center justify-center shrink-0">
              3
            </span>
            <div>
              <span className="font-bold">{t.step3Title}</span>
              <p className="text-[11px] text-slate-500 dark:text-slate-400 hidden md:block">
                {t.step3Desc}
              </p>
            </div>
          </div>
        </div>
      </div>

      {/* Microsoft Word-Like Editor Container */}
      <div className="w-full max-w-5xl bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-2xl shadow-xl overflow-hidden flex flex-col transition-colors">
        {/* Document Title Header Bar */}
        <div className="no-print px-4 sm:px-6 py-2.5 border-b border-slate-200 dark:border-slate-800 flex flex-wrap items-center justify-between gap-3 bg-slate-50/70 dark:bg-slate-950/60">
          <div className="flex-1 min-w-[200px] flex items-center space-x-2 rtl:space-x-reverse">
            <FileText className="h-4 w-4 text-blue-600 shrink-0" />
            <input
              type="text"
              value={documentTitle}
              onChange={e => onChangeTitle(e.target.value)}
              placeholder="Document Title (e.g., Quantum Mechanics Research)"
              className="w-full font-bold text-slate-900 dark:text-white bg-transparent text-sm sm:text-base focus:outline-none focus:border-b-2 focus:border-blue-500 placeholder-slate-400"
            />
          </div>

          {/* Quick Document Status Badges */}
          <div className="flex items-center gap-2 text-xs text-slate-500 dark:text-slate-400">
            <span className="hidden sm:inline">
              <strong>{words}</strong> words
            </span>
            <span className="hidden sm:inline">•</span>
            <span>
              <strong>{totalPages}</strong> {totalPages === 1 ? 'Page' : 'Pages'}
            </span>
          </div>
        </div>

        {/* Word Ribbon Toolbar */}
        <WordRibbon
          formatting={formatting}
          onChangeFormatting={updateFormatting}
          viewMode={viewMode}
          onChangeViewMode={setViewMode}
          onWrapSelection={wrapSelection}
          onApplyBlockStyle={applyBlockStyle}
          onExecCommand={handleExecCommand}
          onUndo={handleUndo}
          onRedo={handleRedo}
          canUndo={historyIndex > 0}
          canRedo={historyIndex < history.length - 1}
          onOpenInsertModal={onOpenInsertModal}
          onOpenTableDialog={() => setShowTableDialog(true)}
          onOpenMathDialog={() => setShowMathDialog(true)}
          onOpenCodeDialog={() => setShowCodeDialog(true)}
          onOpenImageModal={() => setShowImageModal(true)}
          onOpenLinkModal={() => setShowLinkModal(true)}
          onOpenSymbolModal={() => setShowSymbolModal(true)}
          onOpenStatsModal={() => setShowStatsModal(true)}
          onOpenFindReplaceModal={() => setShowFindReplaceModal(true)}
          onOpenWatermarkModal={() => setShowWatermarkModal(true)}
          onImportDocx={handleImportDocx}
          onExportDocx={onExportDocx}
          onExportPdf={onExportPdf}
          onCopyGoogleDocs={onCopyGoogleDocs}
          onSaveDoc={onSaveDoc}
          onClearDoc={onClearDoc}
          onLoadTemplate={handleLoadTemplate}
          isExporting={isExporting}
        />

        {/* Visual Ruler (Word-style) */}
        {formatting.showRuler && (viewMode === 'print' || viewMode === 'web') && (
          <div className="pt-2 bg-slate-100 dark:bg-slate-950 border-b border-slate-200 dark:border-slate-800">
            <WordRuler margins={formatting.margins} orientation={formatting.orientation} />
          </div>
        )}

        {/* Editor Canvas Area */}
        <div className="w-full bg-slate-200/80 dark:bg-slate-950 p-2 sm:p-6 md:p-8 flex justify-center overflow-x-auto min-h-[400px] sm:min-h-[700px] lg:min-h-[800px]">
          {/* ===================== VIEW: PRINT LAYOUT (PAGINATED A4) ===================== */}
          {viewMode === 'print' && (
            <div className="flex flex-col items-center gap-6 w-full max-w-[1100px]">
              {markdown.trim() ? (
                <div
                  className={`relative paper-canvas print-area w-full ${paperOrientationWidth} ${paperMarginClasses} rounded-sm bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-800 shadow-xl transition-all`}
                  style={fontFamilyStyle}
                >
                  {/* Page Header (Word Header) */}
                  <div className="pb-3 mb-6 border-b border-slate-200 dark:border-slate-800 flex items-center justify-between text-[11px] text-slate-400 select-none">
                    <span className="truncate max-w-[300px] font-semibold">{documentTitle}</span>
                    <span>
                      {formatting.paperSize.toUpperCase()} •{' '}
                      {formatting.orientation === 'landscape' ? 'Landscape' : 'Portrait'}
                    </span>
                  </div>

                  {/* Watermark Display */}
                  {formatting.watermark && (
                    <div className="absolute inset-0 flex items-center justify-center pointer-events-none select-none overflow-hidden z-0">
                      <span className="text-6xl sm:text-7xl font-black text-slate-200/50 dark:text-slate-800/40 uppercase tracking-widest rotate-[-35deg]">
                        {formatting.watermark}
                      </span>
                    </div>
                  )}

                  {/* Unified Interactive & Editable Paper Content */}
                  <div
                    ref={paperEditableRef}
                    contentEditable={true}
                    suppressContentEditableWarning={true}
                    onInput={handlePaperInput}
                    onBlur={handlePaperBlur}
                    onMouseUp={saveSelection}
                    onKeyUp={saveSelection}
                    dangerouslySetInnerHTML={{ __html: markdownToHtml(markdown) }}
                    className="relative z-10 prose dark:prose-invert max-w-none text-slate-900 dark:text-slate-100 focus:outline-none min-h-[350px] sm:min-h-[600px] lg:min-h-[750px]"
                  />

                  {/* Page Footer (Word Footer) */}
                  <div className="pt-4 mt-12 border-t border-slate-200 dark:border-slate-800 flex items-center justify-between text-[11px] text-slate-400 select-none">
                    <span>MassiveMark Word Document</span>
                    <span className="font-semibold">
                      Page 1 of {totalPages}
                    </span>
                  </div>
                </div>
              ) : (
                /* Empty state prompt */
                <div
                  onClick={onOpenInsertModal}
                  className="paper-canvas w-full max-w-[850px] min-h-[500px] flex flex-col items-center justify-center border-2 border-dashed border-slate-300 dark:border-slate-800 rounded-xl cursor-pointer hover:border-blue-400 group transition p-8 text-center bg-white dark:bg-slate-900 shadow-lg"
                >
                  <div className="h-16 w-16 rounded-2xl bg-blue-50 dark:bg-blue-950/40 text-blue-600 dark:text-blue-400 flex items-center justify-center mb-3 group-hover:scale-110 transition-transform shadow-sm">
                    <FileText className="h-8 w-8" />
                  </div>
                  <h3 className="text-lg font-bold text-slate-800 dark:text-slate-200 mb-1">
                    Empty Document
                  </h3>
                  <p className="text-xs sm:text-sm text-slate-500 dark:text-slate-400 max-w-md mb-4">
                    Paste AI Markdown (Ctrl+K), type directly, or drop any .docx file here to import.
                  </p>
                  <div className="flex gap-2">
                    <span className="px-4 py-2 rounded-xl bg-red-600 group-hover:bg-red-700 text-white font-bold text-xs sm:text-sm shadow-md shadow-red-600/20">
                      {t.insertButton} (Ctrl+K)
                    </span>
                    <button
                      type="button"
                      onClick={e => {
                        e.stopPropagation();
                        handleLoadTemplate('quantum');
                      }}
                      className="px-4 py-2 rounded-xl border border-slate-300 dark:border-slate-700 hover:bg-slate-100 dark:hover:bg-slate-800 text-xs sm:text-sm font-semibold cursor-pointer"
                    >
                      Load Sample Template
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* ===================== VIEW: WEB LAYOUT (CONTINUOUS) ===================== */}
          {viewMode === 'web' && (
            <div
              className={`paper-canvas print-area w-full ${paperOrientationWidth} ${paperMarginClasses} rounded-xl bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-800 shadow-xl transition-all`}
              style={fontFamilyStyle}
            >
              {markdown.trim() ? (
                <div
                  ref={paperEditableRef}
                  contentEditable={true}
                  suppressContentEditableWarning={true}
                  onInput={handlePaperInput}
                  onBlur={handlePaperBlur}
                  onMouseUp={saveSelection}
                  onKeyUp={saveSelection}
                  dangerouslySetInnerHTML={{ __html: markdownToHtml(markdown) }}
                  className="prose dark:prose-invert max-w-none text-slate-900 dark:text-slate-100 focus:outline-none min-h-[350px] sm:min-h-[600px]"
                />
              ) : (
                <div className="text-center py-24 text-slate-400">
                  <p>Document is empty. Click to type or use Ribbon tools above to start writing.</p>
                </div>
              )}
            </div>
          )}

          {/* ===================== VIEW: SPLIT VIEW (PAPER + SOURCE) ===================== */}
          {viewMode === 'split' && (
            <div className="w-full max-w-6xl grid grid-cols-1 lg:grid-cols-2 gap-4">
              {/* Left: Interactive Paper */}
              <div
                className="paper-canvas p-6 rounded-xl bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-800 shadow-md overflow-y-auto max-h-[750px]"
                style={fontFamilyStyle}
              >
                <div
                  ref={paperEditableRef}
                  contentEditable={true}
                  suppressContentEditableWarning={true}
                  onInput={handlePaperInput}
                  onBlur={handlePaperBlur}
                  onMouseUp={saveSelection}
                  onKeyUp={saveSelection}
                  dangerouslySetInnerHTML={{ __html: markdownToHtml(markdown) }}
                  className="prose dark:prose-invert max-w-none text-xs sm:text-sm text-slate-900 dark:text-slate-100 focus:outline-none"
                />
              </div>

              {/* Right: Markdown Source Editor */}
              <div className="rounded-xl bg-slate-900 text-slate-100 p-4 border border-slate-800 font-mono text-xs sm:text-sm shadow-inner flex flex-col">
                <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider mb-2 block">
                  Markdown Source (Live Sync)
                </span>
                <textarea
                  ref={splitTextareaRef}
                  value={markdown}
                  onChange={e => updateMarkdownWithHistory(e.target.value)}
                  onMouseUp={saveSelection}
                  onKeyUp={saveSelection}
                  rows={30}
                  className="w-full flex-1 bg-transparent text-slate-200 font-mono leading-relaxed focus:outline-none resize-none"
                  placeholder="Enter or edit markdown code here..."
                />
              </div>
            </div>
          )}

          {/* ===================== VIEW: SOURCE CODE ONLY ===================== */}
          {viewMode === 'source' && (
            <div className="w-full max-w-[850px] rounded-xl bg-slate-900 text-slate-100 p-6 border border-slate-800 font-mono text-xs sm:text-sm shadow-inner">
              <div className="flex items-center justify-between pb-3 mb-3 border-b border-slate-800 text-xs text-slate-400">
                <span>Markdown Raw Source</span>
                <span>{linesCount(markdown)} Lines</span>
              </div>
              <textarea
                ref={textareaRef}
                value={markdown}
                onChange={e => updateMarkdownWithHistory(e.target.value)}
                onMouseUp={saveSelection}
                onKeyUp={saveSelection}
                rows={30}
                className="w-full bg-transparent text-slate-200 font-mono leading-relaxed focus:outline-none resize-none"
                placeholder={t.editorPlaceholder}
              />
            </div>
          )}
        </div>

        {/* Word Bottom Status Bar */}
        <div className="no-print px-4 sm:px-6 py-2 border-t border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-950/80 flex flex-wrap items-center justify-between text-xs text-slate-500 dark:text-slate-400">
          {/* Left stats: Page, Words, Read Time */}
          <div className="flex items-center space-x-4 rtl:space-x-reverse">
            <span className="font-semibold text-slate-700 dark:text-slate-300">
              Page 1 of {totalPages}
            </span>

            <button
              type="button"
              onClick={() => setShowStatsModal(true)}
              className="hover:underline flex items-center gap-1 font-semibold text-blue-600 dark:text-blue-400 cursor-pointer"
              title="Click for full document word statistics"
            >
              <span>{words} words</span>
            </button>

            <span className="hidden sm:inline">{characters} characters</span>

            <span className="hidden md:inline">~{readingMinutes} min read</span>
          </div>

          {/* Right controls: View modes & Zoom slider */}
          <div className="flex items-center space-x-3 rtl:space-x-reverse">
            {/* View Mode icons */}
            <div className="flex items-center space-x-1 rtl:space-x-reverse bg-slate-200/60 dark:bg-slate-800/60 p-0.5 rounded-md">
              <button
                type="button"
                onClick={() => setViewMode('print')}
                className={`p-1 rounded cursor-pointer transition ${
                  viewMode === 'print' ? 'bg-white dark:bg-slate-900 text-blue-600 shadow-2xs font-bold' : ''
                }`}
                title="Print Layout"
              >
                <Eye className="h-3.5 w-3.5" />
              </button>
              <button
                type="button"
                onClick={() => setViewMode('web')}
                className={`p-1 rounded cursor-pointer transition ${
                  viewMode === 'web' ? 'bg-white dark:bg-slate-900 text-blue-600 shadow-2xs font-bold' : ''
                }`}
                title="Web Layout"
              >
                <Layers className="h-3.5 w-3.5" />
              </button>
              <button
                type="button"
                onClick={() => setViewMode('split')}
                className={`p-1 rounded cursor-pointer transition ${
                  viewMode === 'split' ? 'bg-white dark:bg-slate-900 text-blue-600 shadow-2xs font-bold' : ''
                }`}
                title="Split View"
              >
                <Split className="h-3.5 w-3.5" />
              </button>
              <button
                type="button"
                onClick={() => setViewMode('source')}
                className={`p-1 rounded cursor-pointer transition ${
                  viewMode === 'source' ? 'bg-white dark:bg-slate-900 text-blue-600 shadow-2xs font-bold' : ''
                }`}
                title="Source Code"
              >
                <FileCode2 className="h-3.5 w-3.5" />
              </button>
            </div>

            <div className="h-3.5 w-px bg-slate-300 dark:bg-slate-700" />

            {/* Zoom controls */}
            <div className="flex items-center space-x-1.5 rtl:space-x-reverse text-xs">
              <button
                type="button"
                onClick={() =>
                  updateFormatting({ zoomLevel: Math.max(75, formatting.zoomLevel - 10) })
                }
                className="hover:text-slate-900 dark:hover:text-white cursor-pointer"
                title="Zoom Out"
              >
                <ZoomOut className="h-3.5 w-3.5" />
              </button>
              <span className="font-mono font-semibold w-10 text-center">
                {formatting.zoomLevel}%
              </span>
              <button
                type="button"
                onClick={() =>
                  updateFormatting({ zoomLevel: Math.min(150, formatting.zoomLevel + 10) })
                }
                className="hover:text-slate-900 dark:hover:text-white cursor-pointer"
                title="Zoom In"
              >
                <ZoomIn className="h-3.5 w-3.5" />
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* ===================== MODALS ===================== */}

      {/* Document Statistics Modal */}
      <DocumentStatsModal
        isOpen={showStatsModal}
        onClose={() => setShowStatsModal(false)}
        markdown={markdown}
        documentTitle={documentTitle}
      />

      {/* Find & Replace Modal */}
      <FindReplaceModal
        isOpen={showFindReplaceModal}
        onClose={() => setShowFindReplaceModal(false)}
        markdown={markdown}
        onReplaceMarkdown={updateMarkdownWithHistory}
      />

      {/* Insert Image Modal */}
      <InsertImageModal
        isOpen={showImageModal}
        onClose={() => setShowImageModal(false)}
        onInsertImage={wrapSelection}
      />

      {/* Insert Link Modal */}
      <InsertLinkModal
        isOpen={showLinkModal}
        onClose={() => setShowLinkModal(false)}
        onInsertLink={wrapSelection}
      />

      {/* Special Symbols Palette */}
      <SymbolPaletteModal
        isOpen={showSymbolModal}
        onClose={() => setShowSymbolModal(false)}
        onInsertSymbol={sym => wrapSelection(sym, '')}
      />

      {/* Watermark Configuration Modal */}
      <WatermarkModal
        isOpen={showWatermarkModal}
        onClose={() => setShowWatermarkModal(false)}
        watermark={formatting.watermark}
        onSetWatermark={wm => updateFormatting({ watermark: wm })}
      />

      {/* Insert Math Dialog */}
      {showMathDialog && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/60 backdrop-blur-sm animate-fadeIn">
          <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-2xl p-6 max-w-lg w-full shadow-2xl space-y-4">
            <h3 className="font-bold text-slate-900 dark:text-white text-base flex items-center gap-2">
              <Sigma className="h-5 w-5 text-purple-600" />
              Word Equation (LaTeX to OMML)
            </h3>

            <div>
              <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300 mb-1">
                LaTeX Formula
              </label>
              <textarea
                value={mathInput}
                onChange={e => setMathInput(e.target.value)}
                rows={3}
                className="w-full p-3 font-mono text-xs sm:text-sm rounded-lg border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-950 text-slate-900 dark:text-white focus:outline-none focus:ring-2 focus:ring-purple-500"
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300 mb-1">
                Equation Live Preview
              </label>
              <div
                className="p-4 bg-slate-50 dark:bg-slate-950 rounded-lg border border-slate-200 dark:border-slate-800 text-center overflow-x-auto min-h-[56px] flex items-center justify-center"
                dangerouslySetInnerHTML={{ __html: renderLatexHtml(mathInput, true) }}
              />
            </div>

            {/* Quick Math Shortcuts */}
            <div className="flex flex-wrap gap-1.5 text-xs">
              {[
                { label: 'Fraction', val: '\\frac{a}{b}' },
                { label: 'Square Root', val: '\\sqrt{x}' },
                { label: 'Sum', val: '\\sum_{i=1}^n x_i' },
                { label: 'Integral', val: '\\int_a^b f(x)dx' },
                { label: 'Einstein', val: 'E = mc^2' },
                { label: 'Quadratic', val: 'x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}' },
              ].map(s => (
                <button
                  key={s.label}
                  type="button"
                  onClick={() => setMathInput(s.val)}
                  className="px-2 py-1 rounded bg-slate-100 dark:bg-slate-800 hover:bg-purple-50 dark:hover:bg-purple-950/40 text-slate-700 dark:text-slate-300 hover:text-purple-600 border border-slate-200 dark:border-slate-700 transition cursor-pointer"
                >
                  {s.label}
                </button>
              ))}
            </div>

            <div className="flex justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setShowMathDialog(false)}
                className="px-4 py-2 rounded-lg border border-slate-300 dark:border-slate-700 text-xs font-semibold cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleInsertMath}
                className="px-4 py-2 rounded-lg bg-purple-600 hover:bg-purple-700 text-white text-xs font-bold shadow-sm cursor-pointer"
              >
                Insert Word Equation
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Insert Table Dialog */}
      {showTableDialog && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/60 backdrop-blur-sm animate-fadeIn">
          <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-2xl p-6 max-w-sm w-full shadow-2xl space-y-4">
            <h3 className="font-bold text-slate-900 dark:text-white text-base flex items-center gap-2">
              <TableIcon className="h-5 w-5 text-blue-600" />
              Custom Table
            </h3>

            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300 mb-1">
                  Rows
                </label>
                <input
                  type="number"
                  min={1}
                  max={25}
                  value={tableRowsCount}
                  onChange={e => setTableRowsCount(parseInt(e.target.value) || 1)}
                  className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-950 text-slate-900 dark:text-white text-sm"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300 mb-1">
                  Columns
                </label>
                <input
                  type="number"
                  min={1}
                  max={12}
                  value={tableColsCount}
                  onChange={e => setTableColsCount(parseInt(e.target.value) || 1)}
                  className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-950 text-slate-900 dark:text-white text-sm"
                />
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setShowTableDialog(false)}
                className="px-4 py-2 rounded-lg border border-slate-300 dark:border-slate-700 text-xs font-semibold cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleInsertTable}
                className="px-4 py-2 rounded-lg bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold shadow-sm cursor-pointer"
              >
                Insert Table
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Insert Code Dialog */}
      {showCodeDialog && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/60 backdrop-blur-sm animate-fadeIn">
          <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-2xl p-6 max-w-md w-full shadow-2xl space-y-4">
            <h3 className="font-bold text-slate-900 dark:text-white text-base flex items-center gap-2">
              <Code className="h-5 w-5 text-emerald-600" />
              Code Block
            </h3>

            <div>
              <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300 mb-1">
                Language
              </label>
              <select
                value={codeLang}
                onChange={e => setCodeLang(e.target.value)}
                className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-950 text-slate-900 dark:text-white text-sm cursor-pointer"
              >
                <option value="python">Python</option>
                <option value="javascript">JavaScript</option>
                <option value="typescript">TypeScript</option>
                <option value="sql">SQL</option>
                <option value="bash">Bash / Shell</option>
                <option value="json">JSON</option>
                <option value="cpp">C++</option>
                <option value="html">HTML / CSS</option>
              </select>
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300 mb-1">
                Code Content
              </label>
              <textarea
                value={codeSnippet}
                onChange={e => setCodeSnippet(e.target.value)}
                rows={6}
                className="w-full p-3 font-mono text-xs rounded-lg border border-slate-300 dark:border-slate-700 bg-slate-950 text-slate-100 focus:outline-none focus:ring-2 focus:ring-emerald-500"
              />
            </div>

            <div className="flex justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setShowCodeDialog(false)}
                className="px-4 py-2 rounded-lg border border-slate-300 dark:border-slate-700 text-xs font-semibold cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleInsertCode}
                className="px-4 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold shadow-sm cursor-pointer"
              >
                Insert Code
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

function linesCount(str: string): number {
  return str.split('\n').length;
}
