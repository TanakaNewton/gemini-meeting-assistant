import React, { useState, useCallback, useEffect, useMemo } from 'react';
import {
  Container,
  Typography,
  TextField,
  Button,
  Box,
  CircularProgress,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Paper,
  FormControl,
  InputLabel,
  Select,
  MenuItem,
  ClickAwayListener,
  Divider,
  Grid,
  Alert,
  IconButton,
  Card,
  CardContent,
  CardActions,
  Chip,
  List,
  ListItem,
  ListItemIcon,
  ListItemText,
  Menu,
  Tooltip,
  Link
} from '@mui/material';
import ClearIcon from '@mui/icons-material/Clear';
import CloudUploadIcon from '@mui/icons-material/CloudUpload';
import DownloadIcon from '@mui/icons-material/Download';
import SummarizeIcon from '@mui/icons-material/Summarize';
import SaveAltIcon from '@mui/icons-material/SaveAlt';
import TopicIcon from '@mui/icons-material/Topic';
import AssignmentTurnedInIcon from '@mui/icons-material/AssignmentTurnedIn';
import DeleteIcon from '@mui/icons-material/Delete';
import AddCircleOutlineIcon from '@mui/icons-material/AddCircleOutline';
import { GoogleGenerativeAI } from "@google/generative-ai";
import Papa from 'papaparse';
import { useDropzone } from 'react-dropzone';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';

// ヘルパー関数: FileオブジェクトをGoogle AI SDKのPart形式に変換
async function fileToGenerativePart(file, onProgress) {
  onProgress('音声ファイルを処理中...');
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      try {
        const base64Data = reader.result.split(',')[1];
        if (!base64Data) {
          throw new Error('ファイルの読み込みに失敗しました。');
        }
        onProgress('音声ファイルの処理完了');
        resolve({
          inlineData: {
            data: base64Data,
            mimeType: file.type
          },
        });
      } catch (error) {
        console.error("Error reading file:", error);
        reject(error);
      }
    };
    reader.onerror = (error) => {
      console.error("FileReader error:", error);
      reject(error);
    };
    reader.readAsDataURL(file);
  });
}

// ヘルパー関数: APIレスポンステキストを整形し、デフォルト話者名と表示用話者名を割り当てる
function parseTranscriptionResponse(responseText) {
  const lines = responseText.trim().split('\n');
  const parsedResults = [];
  const speakerMap = new Map();
  let speakerCounter = 0;
  let uuidCounter = 0;

  lines.forEach((line) => {
    const trimmedLine = line.trim();
    if (trimmedLine) {
      let originalSpeaker = '不明';
      let defaultSpeaker = '';
      const match = trimmedLine.match(/^(.+?):\s*(.*)$/);
      let textContent = trimmedLine;

      if (match && match[1] && match[2]) {
        originalSpeaker = match[1].trim();
        textContent = match[2].trim();
      } else {
        console.warn("Could not parse speaker from line:", trimmedLine);
      }

      if (speakerMap.has(originalSpeaker)) {
        defaultSpeaker = speakerMap.get(originalSpeaker);
      } else {
        defaultSpeaker = `話者${String.fromCharCode(65 + speakerCounter)}`;
        speakerMap.set(originalSpeaker, defaultSpeaker);
        speakerCounter++;
      }

      parsedResults.push({
        id: crypto.randomUUID ? crypto.randomUUID() : `temp-id-${uuidCounter++}`,
        originalSpeaker: originalSpeaker,
        defaultSpeaker: defaultSpeaker,
        displaySpeaker: defaultSpeaker,
        text: textContent,
      });
    }
  });

  if (parsedResults.length === 0 && responseText.trim()) {
    console.warn("Could not parse the response into lines. Displaying raw text.");
    const defaultSpeaker = `話者${String.fromCharCode(65 + speakerCounter)}`;
    return [{
      id: crypto.randomUUID ? crypto.randomUUID() : `temp-id-${uuidCounter++}`,
      originalSpeaker: 'システム応答',
      defaultSpeaker: defaultSpeaker,
      displaySpeaker: defaultSpeaker,
      text: responseText.trim()
    }];
  }
  return parsedResults;
}

const availableModels = [
  { id: "gemini-1.5-pro-latest", name: "Gemini 1.5 Pro (Latest)" },
  { id: "gemini-1.5-flash-latest", name: "Gemini 1.5 Flash (Latest)" },
  { id: "gemini-2.5-pro-exp-03-25", name: "Gemini 2.5 Pro (Experimental)" },
];

// --- ReactMarkdown用のMUIコンポーネントマッピング ---
const markdownComponents = {
  p: (props) => <Typography variant="body1" gutterBottom {...props} />,
  h1: (props) => <Typography variant="h4" gutterBottom {...props} />,
  h2: (props) => <Typography variant="h5" gutterBottom {...props} />,
  h3: (props) => <Typography variant="h6" gutterBottom {...props} />,
  h4: (props) => <Typography variant="subtitle1" gutterBottom {...props} />,
  h5: (props) => <Typography variant="subtitle2" gutterBottom {...props} />,
  h6: (props) => <Typography variant="caption" display="block" gutterBottom {...props} />,
  ul: (props) => <List dense sx={{ paddingLeft: 2, listStyleType: 'disc', pl: 4 }} {...props} />, // デフォルトの黒丸
  ol: (props) => <List dense component="ol" sx={{ paddingLeft: 2, listStyleType: 'decimal', pl: 4 }} {...props} />, // 番号付きリスト
  li: (props) => (
    // ListItemのデフォルトスタイルを無効化し、リストマーカーを表示させる
    <Typography component="li" sx={{ display: 'list-item', paddingLeft: 0, paddingY: 0.2, listStylePosition: 'outside' }} {...props} />
  ),
  a: (props) => <Link {...props} />,
  strong: (props) => <Typography component="span" sx={{ fontWeight: 'bold' }} {...props} />,
  em: (props) => <Typography component="span" sx={{ fontStyle: 'italic' }} {...props} />,
  // テーブル (remark-gfmが必要)
  table: (props) => <TableContainer component={Paper} sx={{ my: 1}}><Table size="small" {...props} /></TableContainer>,
  thead: (props) => <TableHead {...props} />,
  tbody: (props) => <TableBody {...props} />,
  tr: (props) => <TableRow {...props} />,
  th: (props) => <TableCell sx={{ fontWeight: 'bold' }} {...props} />,
  td: (props) => <TableCell {...props} />,
  code: ({node, inline, className, children, ...props}) => {
    const match = /language-(\w+)/.exec(className || '')
    return !inline ? ( // ブロックコード
      <Box sx={{ my: 1, p: 1.5, bgcolor: 'grey.100', borderRadius: 1, overflowX: 'auto', fontFamily: 'monospace', fontSize: '0.875rem' }}>
        <pre style={{ margin: 0 }}><code className={className} {...props}>{children}</code></pre>
      </Box>
    ) : ( // インラインコード
      <Typography component="code" sx={{ px: 0.5, py: 0.2, bgcolor: 'grey.100', borderRadius: 0.5, fontFamily: 'monospace', fontSize: '0.875rem' }} {...props}>
        {children}
      </Typography>
    )
  },
  hr: (props) => <Divider sx={{ my: 2 }} {...props} />,
  blockquote: (props) => <Box component="blockquote" sx={{ borderLeft: 4, borderColor: 'divider', pl: 2, my: 1, fontStyle: 'italic', color: 'text.secondary' }} {...props} />,
};
// --- マッピング定義ここまで ---

function App() {
  const [apiKey, setApiKey] = useState('');
  const [selectedModel, setSelectedModel] = useState("gemini-2.5-pro-exp-03-25");
  const [selectedFile, setSelectedFile] = useState(null);
  const [isLoading, setIsLoading] = useState(false);
  const [results, setResults] = useState([]);
  const [error, setError] = useState('');
  const [progressMessage, setProgressMessage] = useState('');
  const [speakerCount, setSpeakerCount] = useState('');
  const [editingRowId, setEditingRowId] = useState(null);
  const [editingField, setEditingField] = useState(null);
  const [editingValue, setEditingValue] = useState('');
  const [bulkEditNames, setBulkEditNames] = useState({});
  const [summary, setSummary] = useState(null);
  const [isSummarizing, setIsSummarizing] = useState(false);
  const [summaryError, setSummaryError] = useState(null);
  const [keywords, setKeywords] = useState([]);
  const [isExtractingKeywords, setIsExtractingKeywords] = useState(false);
  const [keywordsError, setKeywordsError] = useState(null);
  const [actionItems, setActionItems] = useState([]);
  const [isExtractingActionItems, setIsExtractingActionItems] = useState(false);
  const [actionItemsError, setActionItemsError] = useState(null);
  const [exportMenuAnchorEl, setExportMenuAnchorEl] = useState(null);

  useEffect(() => {
    const uniqueDefaultSpeakers = [...new Set(results.map(r => r.defaultSpeaker))].filter(Boolean);
    setBulkEditNames(prev => {
      const newBulkNames = { ...prev };
      uniqueDefaultSpeakers.forEach(speaker => {
        if (!(speaker in newBulkNames)) {
          newBulkNames[speaker] = '';
        }
      });
      Object.keys(newBulkNames).forEach(key => {
        if (!uniqueDefaultSpeakers.includes(key)) {
          delete newBulkNames[key];
        }
      });
      return newBulkNames;
    });
  }, [results]);

  const uniqueDefaultSpeakersForBulkEdit = useMemo(() => {
    return [...new Set(results.map(r => r.defaultSpeaker))].filter(Boolean).sort();
  }, [results]);

  const handleModelChange = (event) => {
    setSelectedModel(event.target.value);
  };

  const handleSpeakerCountChange = (event) => {
    const value = event.target.value;
    if (value === '' || /^[1-9]\d*$/.test(value)) {
      setSpeakerCount(value);
    }
  };

  const onDrop = useCallback((acceptedFiles) => {
    if (acceptedFiles && acceptedFiles.length > 0) {
      const file = acceptedFiles[0];
      if (file.type.startsWith('audio/')) {
        setSelectedFile(file);
        setError('');
        setResults([]);
        setEditingRowId(null);
        setBulkEditNames({});
        setProgressMessage('');
        setSummary(null);
        setIsSummarizing(false);
        setSummaryError(null);
        setKeywords([]);
        setIsExtractingKeywords(false);
        setKeywordsError(null);
        setActionItems([]);
        setIsExtractingActionItems(false);
        setActionItemsError(null);
      } else {
        setSelectedFile(null);
        setError('音声ファイルを選択してください。');
      }
    }
  }, []);

  const { getRootProps, getInputProps, isDragActive, open } = useDropzone({
    onDrop,
    accept: { 'audio/*': [] },
    noClick: true,
    noKeyboard: true,
    multiple: false
  });

  const handleFileClear = () => {
    setSelectedFile(null);
    setError('');
    setResults([]);
    setEditingRowId(null);
    setBulkEditNames({});
    setProgressMessage('');
    setSpeakerCount('');
    setSummary(null);
    setIsSummarizing(false);
    setSummaryError(null);
    setKeywords([]);
    setIsExtractingKeywords(false);
    setKeywordsError(null);
    setActionItems([]);
    setIsExtractingActionItems(false);
    setActionItemsError(null);
  };

  const handleTranscribe = async () => {
    if (!apiKey || !selectedFile || !selectedModel) {
      setError('APIキー、モデル、音声ファイルを選択してください。');
      return;
    }
    setIsLoading(true);
    setError('');
    setResults([]);
    setEditingRowId(null);
    setBulkEditNames({});
    setSummary(null);
    setIsSummarizing(false);
    setSummaryError(null);
    setKeywords([]);
    setIsExtractingKeywords(false);
    setKeywordsError(null);
    setActionItems([]);
    setIsExtractingActionItems(false);
    setActionItemsError(null);
    setProgressMessage('処理を開始します...');

    try {
      const audioPart = await fileToGenerativePart(selectedFile, setProgressMessage);
      setProgressMessage('APIクライアントを初期化中...');
      const genAI = new GoogleGenerativeAI(apiKey);
      const model = genAI.getGenerativeModel({ model: selectedModel });

      let prompt = `
        以下の音声ファイルを文字起こしし、話者ごとに発言を分けてください。
        出力形式は以下の例のように、各行を "話者X: 発言内容" の形式にしてください。

        例:
        話者A: こんにちは。
        話者B: 今日は良い天気ですね。
        話者A: そうですね、どこかへ出かけたい気分です。
      `;
      const numSpeakers = parseInt(speakerCount, 10);
      if (!isNaN(numSpeakers) && numSpeakers > 0) {
        prompt += `\n\n注記: この会話には${numSpeakers}人の話者が参加しています。これを考慮して話者を区別してください。`;
      }

      setProgressMessage(`Gemini API (${selectedModel}) にリクエスト送信中...`);
      const result = await model.generateContent([prompt, audioPart]);

      setProgressMessage('APIからの応答を受信しました');
      const response = result.response;
      const text = response.text();

      setProgressMessage('応答を解析中...');
      const parsedResults = parseTranscriptionResponse(text);
      setResults(parsedResults);
      setProgressMessage('完了');

    } catch (err) {
      console.error("Full error object:", err);
      let errorMessage = "文字起こし中に予期せぬエラーが発生しました。";
      const statusCode = err?.code || err?.response?.status || err?.response?.error?.code;
      const errorDetails = err?.details || err?.response?.error?.details;

      if (statusCode === 429 || err?.message?.includes('quota') || err?.message?.includes('RESOURCE_EXHAUSTED')) {
        let isModelSpecificQuota = false;
        let violatedModel = '';
        if (Array.isArray(errorDetails)) {
          const quotaFailure = errorDetails.find(detail => detail['@type'] === 'type.googleapis.com/google.rpc.QuotaFailure');
          if (quotaFailure && Array.isArray(quotaFailure.violations)) {
            const modelViolation = quotaFailure.violations.find(v => v?.quotaDimensions?.model === selectedModel);
            if (modelViolation) {
              isModelSpecificQuota = true;
              violatedModel = selectedModel;
            }
          }
        }
        if (isModelSpecificQuota) {
          const modelDisplayName = availableModels.find(m => m.id === violatedModel)?.name || violatedModel;
          errorMessage = `選択されたモデル (${modelDisplayName}) の無料利用枠の上限に達したか、一時的に利用が制限されています。他のモデルを試すか、時間をおいて再度お試しください。`;
        } else {
          errorMessage = 'APIの利用上限(クォータ)に達した可能性があります。プランを確認するか、時間をおいて再度お試しください。';
        }
      } else if (statusCode === 400 && err?.message?.includes('API key not valid')) {
        errorMessage = 'APIキーが無効、または選択したモデルへのアクセス権がありません。キーとモデルの組み合わせを確認してください。';
      } else if (statusCode === 403) {
         errorMessage = '選択したモデルへのアクセスが拒否されました。APIキーに必要な権限が付与されているか確認してください。';
      } else if (statusCode === 404 && err?.message?.includes('Model not found')) {
         errorMessage = `選択されたモデル (${selectedModel}) が見つかりません。`;
      } else if (err?.message?.includes('SAFETY')) {
        errorMessage = 'コンテンツがセーフティポリシーによりブロックされました。';
      } else if (err?.message?.includes('Unsupported audio format')) {
        errorMessage = 'サポートされていない音声ファイル形式です。';
      } else if (err instanceof Error) {
        errorMessage = `エラーが発生しました: ${err.message}`;
      } else if (typeof err === 'string') {
        errorMessage = `エラーが発生しました: ${err}`;
      }
      if (errorDetails) {
        console.error("Error details:", errorDetails);
      }

      setError(errorMessage);
      setResults([]);
      setProgressMessage('エラーが発生しました');
    }
    finally {
      setIsLoading(false);
    }
  };

  const handleCellClick = (rowId, field, currentValue) => {
    if (isLoading || isSummarizing || isExtractingKeywords || isExtractingActionItems) return;
    setEditingRowId(rowId);
    setEditingField(field);
    setEditingValue(currentValue);
  };

  const handleValueSave = useCallback((rowId) => {
    if (editingRowId !== rowId || !editingField) return;

    const newValue = editingValue;
    setResults(prevResults =>
      prevResults.map((row) => {
        if (row.id === rowId) {
          if (editingField === 'speaker') {
            return { ...row, displaySpeaker: newValue.trim() || row.defaultSpeaker || '不明' };
          } else if (editingField === 'text') {
            return { ...row, text: newValue };
          }
        }
        return row;
      })
    );

    setEditingRowId(null);
    setEditingField(null);
    setEditingValue('');
  }, [editingRowId, editingField, editingValue, setResults]);

  const handleKeyDown = (event, rowId) => {
    if (event.key === 'Enter') {
      if (editingField === 'text' && event.shiftKey) {
        return;
      }
      handleValueSave(rowId);
      event.preventDefault();
    } else if (event.key === 'Escape') {
      setEditingRowId(null);
      setEditingField(null);
      setEditingValue('');
    }
  };

  const handleBulkNameChange = (event, defaultSpeaker) => {
    const { value } = event.target;
    setBulkEditNames(prev => ({
      ...prev,
      [defaultSpeaker]: value
    }));
  };

  const handleBulkUpdate = () => {
    setResults(prevResults =>
      prevResults.map(row => {
        const newName = bulkEditNames[row.defaultSpeaker]?.trim();
        return newName ? { ...row, displaySpeaker: newName } : row;
      })
    );
  };

  const handleAddRow = (index) => {
    const newRow = {
      id: crypto.randomUUID ? crypto.randomUUID() : `temp-id-${Date.now()}`,
      originalSpeaker: '',
      defaultSpeaker: '',
      displaySpeaker: '',
      text: '',
    };
    setResults(prevResults => {
      const newResults = [...prevResults];
      newResults.splice(index, 0, newRow);
      return newResults;
    });
  };

  const handleDeleteRow = (idToDelete) => {
    setResults(prevResults => prevResults.filter(row => row.id !== idToDelete));
    if (editingRowId === idToDelete) {
        setEditingRowId(null);
        setEditingField(null);
        setEditingValue('');
    }
  };

  const downloadFile = (content, filename, mimeType) => {
    const blob = new Blob([content], { type: mimeType });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', filename);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 100);
  };

  const getTimestamp = () => {
    const now = new Date();
    return `${now.getFullYear()}${(now.getMonth() + 1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}_${now.getHours().toString().padStart(2, '0')}${now.getMinutes().toString().padStart(2, '0')}${now.getSeconds().toString().padStart(2, '0')}`;
  };

  const handleDownloadCsv = () => {
    if (results.length === 0) {
      handleExportMenuClose();
      return;
    }
    const csvData = results.map(row => ({
      '話者': row.displaySpeaker || '不明',
      '発言内容': row.text
    }));
    const csvString = Papa.unparse(csvData, { header: true, quotes: true, newline: "\r\n" });
    const bom = new Uint8Array([0xEF, 0xBB, 0xBF]);
    const content = new Blob([bom, csvString]);
    downloadFile(content, `transcription_${getTimestamp()}.csv`, 'text/csv;charset=utf-8;');
    handleExportMenuClose();
  };

  const handleDownloadMarkdownTranscription = () => {
    if (results.length === 0) {
        handleExportMenuClose();
        return;
    }
    const markdownString = results.map(row => `**${row.displaySpeaker || '不明'}:** ${row.text}`).join('\n\n');
    downloadFile(markdownString, `transcription_${getTimestamp()}.md`, 'text/markdown;charset=utf-8;');
    handleExportMenuClose();
  };

  const handleDownloadTxtTranscription = () => {
    if (results.length === 0) {
        handleExportMenuClose();
        return;
    }
    const txtString = results.map(row => `${row.displaySpeaker || '不明'}: ${row.text}`).join('\n\n');
    downloadFile(txtString, `transcription_${getTimestamp()}.txt`, 'text/plain;charset=utf-8;');
    handleExportMenuClose();
  };

  const handleSummarize = async () => {
    if (!results || results.length === 0 || !apiKey) {
      setSummaryError('要約する文字起こし結果またはAPIキーがありません。');
      return;
    }
    setIsSummarizing(true);
    setSummary(null);
    setSummaryError(null);
    setProgressMessage('');

    try {
      const conversationText = results.map(row => `${row.displaySpeaker}: ${row.text}`).join('\n');
      // プロンプトを改善：より詳細なMarkdown形式を期待する
      const summaryPrompt = `以下の会話を簡潔に要約し、見出し、箇条書き、太字などを使用して分かりやすくMarkdown形式で出力してください:\n\n\`\`\`\n${conversationText}\n\`\`\``;

      const genAI = new GoogleGenerativeAI(apiKey);
      const model = genAI.getGenerativeModel({ model: selectedModel });
      const result = await model.generateContent(summaryPrompt);
      const response = result.response;
      let summaryText = response.text();

       // 応答がMarkdownブロックで囲まれている場合、中身だけを取り出す
       const markdownBlockMatch = summaryText.match(/^```markdown\s*([\s\S]*?)\s*```$/);
       if (markdownBlockMatch) {
         summaryText = markdownBlockMatch[1];
       } else {
         // ``` のみで囲まれている場合も考慮
         const codeBlockMatch = summaryText.match(/^```\s*([\s\S]*?)\s*```$/);
         if (codeBlockMatch) {
           summaryText = codeBlockMatch[1];
         }
       }

      setSummary(summaryText.trim()); // 前後の空白を削除

    } catch (err) {
      console.error("Error during summarization:", err);
      let msg = "要約中にエラーが発生しました。";
      if (err instanceof Error) {
        msg += ` 詳細: ${err.message}`;
      }
      setSummaryError(msg);
    } finally {
      setIsSummarizing(false);
    }
  };

  const handleDownloadSummaryMarkdown = () => {
    if (!summary) { return; }
    downloadFile(summary, `summary_${getTimestamp()}.md`, 'text/markdown;charset=utf-8;');
  };

  const handleExtractKeywords = async () => {
    if (!results || results.length === 0 || !apiKey) {
      setKeywordsError('キーワードを抽出する文字起こし結果またはAPIキーがありません。');
      return;
    }
    setIsExtractingKeywords(true);
    setKeywords([]);
    setKeywordsError(null);
    setProgressMessage('');

    try {
      const conversationText = results.map(row => `${row.displaySpeaker}: ${row.text}`).join('\n');
      const prompt = `以下の会話から重要なキーワードやトピックを抽出し、JSON配列の形式で出力してください。各要素は文字列である必要があります。

例:
\`\`\`json
[
  "決済機能設計(山田、田中)",
  "納期遅延の可能性",
  "追加仕様の確認"
]
\`\`\`

会話:
\`\`\`
${conversationText}
\`\`\``;

      const genAI = new GoogleGenerativeAI(apiKey);
      const model = genAI.getGenerativeModel({ model: selectedModel });
      const result = await model.generateContent(prompt);
      const response = result.response;
      const text = response.text();

      let items = [];
      try {
        const match = text.match(/```json\s*([\s\S]*?)\s*```/);
        const jsonStr = match ? match[1] : text;
        items = JSON.parse(jsonStr);
        if (!Array.isArray(items) || (items.length > 0 && typeof items[0] !== 'string')) {
             throw new Error("API応答が期待されたJSON配列(文字列)形式ではありません。");
        }
        items = items.map(kw => kw.trim()).filter(kw => kw.length > 0);
      } catch (e) {
        console.error("Failed to parse keywords JSON:", e, "Raw response text:", text);
        throw new Error(`キーワードリストの解析に失敗しました。応答が期待されるJSON配列形式ではありませんでした。\n詳細: ${e.message}`);
      }
      setKeywords(items);

    } catch (err) {
      console.error("Error during keyword extraction:", err);
      let msg = "キーワード抽出中にエラーが発生しました。";
      if (err instanceof Error) {
        msg += ` 詳細: ${err.message}`;
      } else if (typeof err === 'string') {
        msg += ` 詳細: ${err}`;
      }
      setKeywordsError(msg);
    } finally {
      setIsExtractingKeywords(false);
    }
  };

  const handleExtractActionItems = async () => {
    if (!results || results.length === 0 || !apiKey) {
      setActionItemsError('アクションアイテムを抽出する文字起こし結果またはAPIキーがありません。');
      return;
    }
    setIsExtractingActionItems(true);
    setActionItems([]);
    setActionItemsError(null);
    setProgressMessage('');

    try {
      const conversationText = results.map(row => `${row.displaySpeaker}: ${row.text}`).join('\n');
      const prompt = `以下の会話からアクションアイテム（担当者、タスク内容、期限）を抽出し、以下のJSON形式の配列で出力してください。該当する情報がない場合は省略するかnullを使用してください。期限はYYYY-MM-DD形式で記述してください。

出力形式:
\`\`\`json
[
  { "assignee": "担当者名", "task": "タスク内容", "dueDate": "YYYY-MM-DD または null" },
  ...
]
\`\`\`

会話:
\`\`\`
${conversationText}
\`\`\``;

      const genAI = new GoogleGenerativeAI(apiKey);
      const model = genAI.getGenerativeModel({ model: selectedModel });
      const result = await model.generateContent(prompt);
      const response = result.response;
      const text = response.text();

      let items = [];
      try {
        const match = text.match(/```json\s*([\s\S]*?)\s*```/);
        const jsonStr = match ? match[1] : text;
        items = JSON.parse(jsonStr);
        if (!Array.isArray(items) || (items.length > 0 && (typeof items[0].assignee === 'undefined' || typeof items[0].task === 'undefined'))) {
             throw new Error("API応答が期待されたJSON形式ではありません。");
        }
      } catch (e) {
        console.error("Failed to parse action items JSON:", e, "Raw response text:", text);
        throw new Error(`API応答の解析に失敗しました。\n詳細: ${e.message}`);
      }

      items = items.map(item => ({
        ...item,
        dueDate: item.dueDate || null
      }));

      setActionItems(items);

    } catch (err) {
      console.error("Error during action item extraction:", err);
      let msg = "アクションアイテム抽出中にエラーが発生しました。";
      if (err instanceof Error) {
        msg += ` 詳細: ${err.message}`;
      }
      setActionItemsError(msg);
    } finally {
      setIsExtractingActionItems(false);
    }
  };

  const handleExportMenuOpen = (event) => {
    setExportMenuAnchorEl(event.currentTarget);
  };

  const handleExportMenuClose = () => {
    setExportMenuAnchorEl(null);
  };


  // === レンダリング ===
  return (
    <Container maxWidth="lg">
      <Box sx={{ my: 4 }}>
        <Typography variant="h4" component="h1" gutterBottom>
          Gemini 議事録作成
        </Typography>

        {/* === 上部設定エリア (Grid V2 対応) === */}
        <Grid container spacing={2} sx={{ mb: 3 }} alignItems="flex-start">
          {/* item プロパティは不要 */}
          <Grid xs={12} md={5}>
            <TextField
              label="Gemini API Key"
              variant="outlined"
              fullWidth
              type="password"
              value={apiKey}
              onChange={(e) => setApiKey(e.target.value)}
              disabled={isLoading || isSummarizing || isExtractingKeywords || isExtractingActionItems}
              size="small"
            />
          </Grid>
          <Grid xs={12} md={4}>
            <FormControl fullWidth disabled={isLoading || isSummarizing || isExtractingKeywords || isExtractingActionItems} size="small">
              <InputLabel id="model-select-label">モデルを選択</InputLabel>
              <Select
                labelId="model-select-label"
                id="model-select"
                value={selectedModel}
                label="モデルを選択"
                onChange={handleModelChange}
              >
                {availableModels.map((model) => (
                  <MenuItem key={model.id} value={model.id}>
                    {model.name}
                  </MenuItem>
                ))}
              </Select>
            </FormControl>
          </Grid>
          <Grid xs={6} md={3}>
             <TextField
                label="話者数 (任意)"
                variant="outlined"
                type="number"
                value={speakerCount}
                onChange={handleSpeakerCountChange}
                disabled={isLoading || isSummarizing || isExtractingKeywords || isExtractingActionItems}
                size="small"
                fullWidth
                InputProps={{ inputProps: { min: 1 } }}
                helperText="精度向上に役立つ場合があります"
             />
          </Grid>
          <Grid xs={6} md={12} sx={{ textAlign: { xs: 'left', md: 'right' } }}>
             {results.length > 0 && !isLoading && (
                 <>
                  <Button
                    id="export-button"
                    aria-controls={exportMenuAnchorEl ? 'export-menu' : undefined}
                    aria-haspopup="true"
                    aria-expanded={exportMenuAnchorEl ? 'true' : undefined}
                    variant="outlined"
                    onClick={handleExportMenuOpen}
                    disabled={results.length === 0 || isLoading || isSummarizing || isExtractingKeywords || isExtractingActionItems}
                    size="medium"
                    startIcon={<DownloadIcon />}
                    sx={{ mt: { xs: 1, md: 0 } }}
                  >
                    ダウンロード
                  </Button>
                  <Menu
                    id="export-menu"
                    anchorEl={exportMenuAnchorEl}
                    open={Boolean(exportMenuAnchorEl)}
                    onClose={handleExportMenuClose}
                    MenuListProps={{
                      'aria-labelledby': 'export-button',
                    }}
                  >
                    <MenuItem onClick={handleDownloadCsv}>CSV形式 (.csv)</MenuItem>
                    <MenuItem onClick={handleDownloadMarkdownTranscription}>Markdown形式 (.md)</MenuItem>
                    <MenuItem onClick={handleDownloadTxtTranscription}>テキスト形式 (.txt)</MenuItem>
                  </Menu>
                 </>
             )}
          </Grid>
        </Grid>

        {/* === ファイル選択エリア === */}
        {!selectedFile && (
          <Box
            {...getRootProps()}
            sx={{
              border: '3px dashed',
              borderColor: isDragActive ? 'primary.main' : 'grey.400',
              borderRadius: 2,
              p: 4,
              minHeight: 200,
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              justifyContent: 'center',
              textAlign: 'center',
              backgroundColor: isDragActive ? 'action.hover' : 'grey.50',
              cursor: 'pointer',
              transition: 'border-color 0.2s ease-in-out, background-color 0.2s ease-in-out',
              mb: 3
            }}
          >
            <input {...getInputProps()} />
            <CloudUploadIcon sx={{ fontSize: 60, mb: 2, color: 'grey.500' }} />
            {isDragActive ? (
              <Typography variant="h6">ここにファイルをドロップ</Typography>
            ) : (
              <>
                <Typography variant="h6" gutterBottom>
                  音声ファイルをドラッグ＆ドロップ
                </Typography>
                <Typography color="text.secondary">または</Typography>
                <Button
                  variant="contained"
                  onClick={open}
                  disabled={isLoading || isSummarizing || isExtractingKeywords || isExtractingActionItems}
                  sx={{ mt: 2 }}
                >
                  ファイルを選択
                </Button>
              </>
            )}
          </Box>
        )}

        {/* === 選択中ファイル表示 & 文字起こしボタン === */}
        {selectedFile && (
          <Paper elevation={1} sx={{ p: 2, display: 'flex', alignItems: 'center', justifyContent: 'space-between', mb: 3 }}>
            <Box sx={{ display: 'flex', alignItems: 'center', overflow: 'hidden' }}>
              <Typography sx={{ mr: 1, flexShrink: 0 }}>選択中のファイル:</Typography>
              <Typography noWrap title={selectedFile.name} sx={{ fontWeight: 'medium', mr: 1 }}>
                {selectedFile.name} ({ (selectedFile.size / 1024 / 1024).toFixed(2) } MB)
              </Typography>
              <IconButton
                 onClick={handleFileClear}
                 size="small"
                 sx={{ ml: 'auto' }}
                 title="ファイルをクリア"
                 disabled={isLoading || isSummarizing || isExtractingKeywords || isExtractingActionItems}
              >
                 <ClearIcon fontSize="small"/>
              </IconButton>
            </Box>
            <Button
              variant="contained"
              color="primary"
              onClick={handleTranscribe}
              disabled={!apiKey || !selectedFile || !selectedModel || isLoading || isSummarizing || isExtractingKeywords || isExtractingActionItems}
              size="large"
            >
              文字起こし開始
            </Button>
          </Paper>
        )}

        {/* === ローディングとエラー/完了表示 (トップレベル) === */}
        {isLoading && (
          <Box sx={{ display: 'flex', justifyContent: 'center', alignItems: 'center', my: 2, flexDirection: 'column' }}>
             <Box sx={{ display: 'flex', alignItems: 'center'}}>
                 <CircularProgress size={24} />
                 <Typography sx={{ ml: 1 }}>処理中...</Typography>
             </Box>
             {progressMessage && progressMessage !== '完了' && progressMessage !== 'エラーが発生しました' && (
                <Typography variant="body2" color="text.secondary" sx={{mt: 1}}>
                    {progressMessage}
                </Typography>
             )}
          </Box>
        )}
        {error && !isLoading && (
          <Alert severity="error" sx={{ my: 2 }}>{error}</Alert>
        )}
        {!isLoading && !isSummarizing && !isExtractingKeywords && !isExtractingActionItems && progressMessage === '完了' && results.length > 0 && !error && (
          <Alert severity="success" sx={{ my: 2 }}>文字起こしが完了しました。</Alert>
        )}

        {/* === 結果表示エリア (文字起こし完了後) === */}
        {results.length > 0 && !isLoading && (
          <>
            {/* --- 要約セクション --- */}
            <Card variant="outlined" sx={{ my: 3 }}>
              <CardContent>
                <Typography variant="h6" gutterBottom>
                  要約
                </Typography>
                {isSummarizing && (
                  <Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'center', my: 2 }}>
                    <CircularProgress size={20} sx={{mr: 1}}/>
                    <Typography>要約を生成中...</Typography>
                  </Box>
                )}
                {summaryError && !isSummarizing && (
                  <Alert severity="error" sx={{mb: 1}}>{summaryError}</Alert>
                )}
                {!isSummarizing && !summaryError && summary && (
                  <Alert severity="info" sx={{ mb: 1 }}>要約が生成されました。</Alert>
                )}
                {summary && !isSummarizing && (
                  // Boxに typography: 'body2' がないことを確認
                  <Box sx={{ maxHeight: '300px', overflowY: 'auto', p: 1, border: '1px solid', borderColor: 'divider', borderRadius: 1 }}>
                    {/* components プロパティを追加してMUIコンポーネントにマッピング */}
                    <ReactMarkdown
                      remarkPlugins={[remarkGfm]}
                      components={markdownComponents}
                    >
                      {summary}
                    </ReactMarkdown>
                  </Box>
                )}
                {!summary && !isSummarizing && !summaryError && (
                  <Typography variant="body2" color="text.secondary">
                    要約は作成されていません。
                  </Typography>
                )}
              </CardContent>
              <CardActions sx={{ justifyContent: 'space-between' }}>
                 {summary && !isSummarizing && !summaryError && (
                    <Button
                       size="small"
                       onClick={handleDownloadSummaryMarkdown}
                       startIcon={<SaveAltIcon />}
                       disabled={isLoading || isSummarizing || isExtractingKeywords || isExtractingActionItems}
                       sx={{ textTransform: 'none' }} // 大文字変換を無効化
                    >
                       Markdownで保存
                    </Button>
                 )}
                 {(!summary || isSummarizing || summaryError) ? <Box sx={{ flexGrow: 1 }} /> : null}
                 <Button
                   size="small"
                   onClick={handleSummarize}
                   disabled={isSummarizing || isLoading || isExtractingKeywords || isExtractingActionItems}
                   startIcon={<SummarizeIcon />}
                   sx={{ textTransform: 'none' }} // 大文字変換を無効化 (必要に応じて)
                 >
                   {summary ? '要約を再生成' : '要約を作成'}
                 </Button>
              </CardActions>
            </Card>

             {/* --- キーワード抽出セクション --- */}
             <Card variant="outlined" sx={{ my: 3 }}>
               <CardContent>
                 <Typography variant="h6" gutterBottom>
                   キーワード / トピック
                 </Typography>
                 {isExtractingKeywords && (
                   <Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'center', my: 2 }}>
                     <CircularProgress size={20} sx={{mr: 1}}/>
                     <Typography>キーワードを抽出中...</Typography>
                   </Box>
                 )}
                 {keywordsError && !isExtractingKeywords && (
                   <Alert severity="error" sx={{mb: 1}}>{keywordsError}</Alert>
                 )}
                 {!isExtractingKeywords && !keywordsError && keywords.length > 0 && (
                   <Alert severity="info" sx={{ mb: 1 }}>キーワードが抽出されました。</Alert>
                 )}
                 {keywords.length > 0 && !isExtractingKeywords && !keywordsError && (
                   <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 1 }}>
                     {keywords.map((keyword, index) => (
                       <Chip key={index} label={keyword} size="small" />
                     ))}
                   </Box>
                 )}
                 {!keywordsError && !isExtractingKeywords && keywords.length === 0 && (
                   <Typography variant="body2" color="text.secondary">
                     キーワードは抽出されていません。
                   </Typography>
                 )}
               </CardContent>
               <CardActions sx={{ justifyContent: 'flex-end' }}>
                  <Button
                     size="small"
                     onClick={handleExtractKeywords}
                     disabled={isExtractingKeywords || isLoading || isSummarizing || isExtractingActionItems}
                     startIcon={<TopicIcon />}
                     sx={{ textTransform: 'none' }} // 大文字変換を無効化 (必要に応じて)
                  >
                     {keywords.length > 0 ? 'キーワードを再抽出' : 'キーワードを抽出'}
                  </Button>
               </CardActions>
             </Card>

             {/* --- アクションアイテム抽出セクション --- */}
             <Card variant="outlined" sx={{ my: 3 }}>
               <CardContent>
                 <Typography variant="h6" gutterBottom>
                   アクションアイテム
                 </Typography>
                 {isExtractingActionItems && (
                   <Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'center', my: 2 }}>
                     <CircularProgress size={20} sx={{mr: 1}}/>
                     <Typography>アクションアイテムを抽出中...</Typography>
                   </Box>
                 )}
                 {actionItemsError && !isExtractingActionItems && (
                   <Alert severity="error" sx={{mb: 1}}>{actionItemsError}</Alert>
                 )}
                 {!isExtractingActionItems && !actionItemsError && actionItems.length > 0 && (
                   <Alert severity="info" sx={{ mb: 1 }}>アクションアイテムが抽出されました。</Alert>
                 )}
                 {actionItems.length > 0 && !isExtractingActionItems && !actionItemsError && (
                   <List dense>
                     {actionItems.map((item, index) => (
                       <ListItem key={index} disablePadding>
                         <ListItemIcon sx={{ minWidth: '30px' }}>
                           <AssignmentTurnedInIcon fontSize="small" />
                         </ListItemIcon>
                         <ListItemText
                           primary={item.task || 'タスク不明'}
                           secondary={`担当: ${item.assignee || '未定'} ${item.dueDate ? `| 期限: ${item.dueDate}` : ''}`}
                         />
                       </ListItem>
                     ))}
                   </List>
                 )}
                 {!actionItemsError && !isExtractingActionItems && actionItems.length === 0 && (
                   <Typography variant="body2" color="text.secondary">
                     アクションアイテムは抽出されていません。
                   </Typography>
                 )}
               </CardContent>
               <CardActions sx={{ justifyContent: 'flex-end' }}>
                 <Button
                    size="small"
                    onClick={handleExtractActionItems}
                    disabled={isExtractingActionItems || isLoading || isSummarizing || isExtractingKeywords}
                    startIcon={<AssignmentTurnedInIcon />}
                    sx={{ textTransform: 'none' }} // 大文字変換を無効化 (必要に応じて)
                 >
                   {actionItems.length > 0 ? 'アクションアイテムを再抽出' : 'アクションアイテムを抽出'}
                 </Button>
               </CardActions>
             </Card>


            {/* --- 一括編集セクション (Grid V2 対応) --- */}
            <Box sx={{ my: 3, p: 2, border: '1px solid', borderColor: 'divider', borderRadius: 1 }}>
              <Typography variant="h6" gutterBottom>
                話者名の一括編集
              </Typography>
              <Grid container spacing={2} alignItems="center">
                {uniqueDefaultSpeakersForBulkEdit.map((defaultSpeaker) => (
                  <Grid xs={12} sm={6} md={4} key={defaultSpeaker}>
                     <Box sx={{ display: 'flex', alignItems: 'baseline' }}>
                        <Typography sx={{ mr: 1, fontWeight: 'bold', whiteSpace: 'nowrap' }}>
                           {defaultSpeaker}:
                        </Typography>
                        <TextField
                           variant="standard"
                           size="small"
                           placeholder={`「${defaultSpeaker}」の新しい名前`}
                           value={bulkEditNames[defaultSpeaker] || ''}
                           onChange={(e) => handleBulkNameChange(e, defaultSpeaker)}
                           disabled={isLoading || isSummarizing || isExtractingKeywords || isExtractingActionItems}
                           fullWidth
                        />
                     </Box>
                  </Grid>
                ))}
                <Grid xs={12}>
                     <Button
                        variant="contained"
                        size="small"
                        onClick={handleBulkUpdate}
                        disabled={Object.values(bulkEditNames).every(name => !name?.trim()) || isLoading || isSummarizing || isExtractingKeywords || isExtractingActionItems}
                        sx={{ textTransform: 'none' }} // 大文字変換を無効化 (必要に応じて)
                     >
                        一括変更を適用
                     </Button>
                </Grid>
              </Grid>
            </Box>

            <Divider sx={{ my: 2 }} />

            {/* --- 個別編集テーブル --- */}
            <Typography variant="h6" gutterBottom sx={{ mb: 1 }}>
                文字起こし結果 ({availableModels.find(m => m.id === selectedModel)?.name || selectedModel}) - セルをクリックして編集
            </Typography>
            <TableContainer component={Paper} sx={{ position: 'relative', overflow: 'visible' }}>
              <Table sx={{ minWidth: 750 }} aria-label="transcription results table">
                <TableHead>
                  <TableRow sx={{ position: 'relative' }}>
                    <TableCell sx={{ fontWeight: 'bold', width: '15%' }}>話者</TableCell>
                    <TableCell sx={{ fontWeight: 'bold' }}>発言内容</TableCell>
                    {/* ヘッダーテキスト削除 */}
                    <TableCell align="center" sx={{ width: '60px', position: 'relative' }}>
                        {/* ヘッダー右下に行頭追加ボタンを配置 */}
                        <Tooltip title="ここに行を追加" placement="left">
                            <IconButton
                                size="small"
                                onClick={() => handleAddRow(0)} // 先頭に追加
                                disabled={isLoading || isSummarizing || isExtractingKeywords || isExtractingActionItems}
                                sx={{
                                    position: 'absolute',
                                    bottom: 0, // ヘッダー行の下端
                                    right: 0, // セルの右端
                                    transform: 'translate(50%, 50%)', // 右下に50%ずらす
                                    zIndex: 20,
                                    backgroundColor: 'background.paper',
                                    border: '1px solid',
                                    borderColor: 'divider',
                                    '&:hover': { backgroundColor: 'action.hover' },
                                    pointerEvents: 'auto'
                                }}
                            >
                                <AddCircleOutlineIcon fontSize="small" />
                            </IconButton>
                        </Tooltip>
                    </TableCell>
                  </TableRow>
                </TableHead>
                <TableBody>
                  {results.map((row, index) => {
                    const displayName = row.displaySpeaker;
                    const isEditingThisSpeaker = editingRowId === row.id && editingField === 'speaker';
                    const isEditingThisText = editingRowId === row.id && editingField === 'text';

                    return (
                      // Hydrationエラー回避のため、tr直下に不要な空白を入れない
                      <TableRow key={row.id} sx={{ '&:last-child td, &:last-child th': { border: 0 } }}>
                        <TableCell
                          component="th"
                          scope="row"
                          sx={{
                             whiteSpace: 'pre-wrap',
                             verticalAlign: 'top',
                             width: '15%',
                             cursor: !isEditingThisSpeaker ? 'pointer' : 'default',
                             minHeight: editingRowId === row.id ? 'auto' : '2.5em',
                             height: 'auto'
                          }}
                          onClick={() => !isEditingThisSpeaker && !(isLoading || isSummarizing || isExtractingKeywords || isExtractingActionItems) && handleCellClick(row.id, 'speaker', displayName)}
                        >{isEditingThisSpeaker ? (
                            <ClickAwayListener onClickAway={() => handleValueSave(row.id)}>
                              <TextField
                                value={editingValue}
                                onChange={(e) => setEditingValue(e.target.value)}
                                onKeyDown={(e) => handleKeyDown(e, row.id)}
                                variant="standard"
                                autoFocus
                                fullWidth
                                size="small"
                              />
                            </ClickAwayListener>
                          ) : (
                            displayName || <Box sx={{ minHeight: '1.4375em' }}> </Box>
                          )}</TableCell>
                        <TableCell
                          sx={{
                             whiteSpace: 'pre-wrap',
                             verticalAlign: 'top',
                             cursor: !isEditingThisText ? 'pointer' : 'default',
                             minHeight: editingRowId === row.id ? 'auto' : '2.5em',
                             height: 'auto'
                          }}
                          onClick={() => !isEditingThisText && !(isLoading || isSummarizing || isExtractingKeywords || isExtractingActionItems) && handleCellClick(row.id, 'text', row.text)}
                        >{isEditingThisText ? (
                             <ClickAwayListener onClickAway={() => handleValueSave(row.id)}>
                               <TextField
                                 value={editingValue}
                                 onChange={(e) => setEditingValue(e.target.value)}
                                 onKeyDown={(e) => handleKeyDown(e, row.id)}
                                 variant="outlined"
                                 autoFocus
                                 fullWidth
                                 multiline
                                 minRows={2}
                                 size="small"
                               />
                             </ClickAwayListener>
                           ) : (
                             row.text || <Box sx={{ minHeight: '1.4375em' }}> </Box>
                           )}</TableCell>
                        <TableCell
                          align="center"
                          sx={{
                            position: 'relative',
                            verticalAlign: 'middle',
                            padding: '0 8px',
                            width: '60px',
                            overflow: 'visible'
                          }}
                        >
                           <Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100%' }}>
                               <Tooltip title="この行を削除" placement="left">
                                   <IconButton
                                       size="small"
                                       onClick={() => handleDeleteRow(row.id)}
                                       disabled={isLoading || isSummarizing || isExtractingKeywords || isExtractingActionItems}
                                   >
                                       <DeleteIcon fontSize="inherit" />
                                   </IconButton>
                               </Tooltip>
                           </Box>
                           <Tooltip title="ここに行を追加" placement="left">
                               <IconButton
                                   size="small"
                                   onClick={() => handleAddRow(index + 1)}
                                   disabled={isLoading || isSummarizing || isExtractingKeywords || isExtractingActionItems}
                                   sx={{
                                       position: 'absolute',
                                       bottom: 0,
                                       right: 0,
                                       transform: 'translate(50%, 50%)',
                                       zIndex: 20,
                                       backgroundColor: 'background.paper',
                                       border: '1px solid',
                                       borderColor: 'divider',
                                       '&:hover': { backgroundColor: 'action.hover' },
                                       pointerEvents: 'auto'
                                   }}
                               >
                                   <AddCircleOutlineIcon fontSize="inherit" />
                               </IconButton>
                           </Tooltip></TableCell>
                      </TableRow>
                    );
                  })}
                </TableBody>
              </Table>
            </TableContainer>
          </>
        )}
      </Box>
    </Container>
  );
}

export default App;