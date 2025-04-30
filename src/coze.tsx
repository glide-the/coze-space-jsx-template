import React, { useState } from 'react';
import { Download, FileText, FileInput, FileOutput, FileCheck, FileSearch, FilePlus, FileMinus, FileX, File, FileDigit, FileClock, FileSpreadsheet, FileBarChart2, FileCode2, FileImage, FileAudio2, FileVideo2, FileArchive, FileSignature, FileDiff, FileHeart, FileJson, FileKey2, FileLock2, FilePieChart, FileScan, FileSearch2, FileStack, FileTerminal, FileType2, FileUp, FileVolume2, FileWarning, FileZip, ChevronDown, ChevronRight, Check, X, Plus, Minus, BookOpen, Brain, Cpu, Activity, Heart, Shield } from 'lucide-react';
import { motion } from 'framer-motion';
import * as XLSX from 'xlsx';

type QuestionCategory = '基础知识' | '智力技能' | '认知策略' | '动作技能' | '情感感知';
type QuestionTypes = {
  [K in QuestionCategory]: string[];
};

type GeneratedData = {
  id: string;
  type: string;
  category: QuestionCategory;
  model: string;
  nodes: number;
  attachment: string;
};

const QuestionGenerator = () => {
  const [model, setModel] = useState('DeepseekR1');
  const [questionTypes, setQuestionTypes] = useState<QuestionTypes>({
    '基础知识': ['概念理解'],
    '智力技能': ['数值计算'],
    '认知策略': ['COT'],
    '动作技能': ['RAG'],
    '情感感知': ['情绪']
  });
  const [nodeCount, setNodeCount] = useState(1);
  const [hasAttachment, setHasAttachment] = useState(false);
  const [questionCount, setQuestionCount] = useState(5);
  const [pdfFile, setPdfFile] = useState<File | null>(null);
  const [isGenerating, setIsGenerating] = useState(false);
  const [generatedData, setGeneratedData] = useState<GeneratedData[]>([]);
  const [expandedCategories, setExpandedCategories] = useState({
    '基础知识': true,
    '智力技能': false,
    '认知策略': false,
    '动作技能': false,
    '情感感知': false
  });

  interface FormData {
    name: string;
    email: string;
    message: string;
  }

  const [formData, setFormData] = useState<FormData>({ name: '', email: '', message: '' });
  const [submittedData, setSubmittedData] = useState<FormData | null>(null);

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    const { name, value } = e.target;
    setFormData(prev => ({
      ...prev,
      [name]: value
    }));
  };

  const handleSubmit = (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    setSubmittedData(formData);
  };

  const toggleCategory = (category: QuestionCategory) => {
    setExpandedCategories(prev => ({
      ...prev,
      [category]: !prev[category]
    }));
  };

  const toggleQuestionType = (category: QuestionCategory, type: string) => {
    setQuestionTypes(prev => {
      const currentTypes = [...prev[category]];
      const index = currentTypes.indexOf(type);
      
      if (index === -1) {
        // 选中新选项时，只更新当前分类，保持其他分类不变
        return {
          ...prev,
          [category]: [type]
        };
      } else {
        // 点击已选中的选项时，保持选中状态
        return prev;
      }
    });
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e?.target?.files?.[0];
    if (file) {
      setPdfFile(file);
    }
  };

  const generateQuestions = async () => {
    // 验证所有必填内容
    const selectedCategory = Object.keys(questionTypes).find(cat => questionTypes[cat as QuestionCategory]?.length > 0) as QuestionCategory | undefined;
    if (!selectedCategory) {
      console.log('错误：请至少选择一个题目类型');
      return;
    }

    // 打印提交的内容
    console.log('提交的内容:', {
      model,
      questionTypes,
      nodeCount,
      hasAttachment,
      questionCount,
      pdfFile: pdfFile ? pdfFile.name : null
    });

    setIsGenerating(true);
    
    try {
      const formData = new FormData();
      if (pdfFile) {
        formData.append('file', pdfFile);
      }
      
      // 添加其他数据
      const jsonData = {
        model,
        questionTypes,
        nodeCount,
        hasAttachment,
        questionCount,
        selectedCategory,
        selectedType: questionTypes[selectedCategory]?.[0] || '概念理解'
      };
      formData.append('json', JSON.stringify(jsonData));

      const response = await fetch('http://localhost:8000/extract_standard_file', {
        method: 'POST',
        body: formData,
      });

      const result = await response.json();
      
      if (result.code === 0) {
        // 处理成功响应
        const data = result.data.map((item: any, i: number) => ({
          id: `T${(i + 1).toString().padStart(3, '0')}`,
          type: questionTypes[selectedCategory]?.[0] || '概念理解',
          category: selectedCategory,
          model,
          nodes: nodeCount,
          attachment: JSON.stringify(item)
        }));
        setGeneratedData(data);
      } else {
        console.error('API请求失败:', result.msg);
      }
    } catch (error) {
      console.error('请求出错:', error);
    } finally {
      setIsGenerating(false);
    }
  };

  const downloadExcel = () => {
    if (generatedData.length === 0) return;

    // Convert data to worksheet format
    const worksheet = XLSX.utils.aoa_to_sheet([
      ['题目ID', '分类', '类型', '模型', '节点', '附件'], // Headers
      ...generatedData.map(row => [
        row.id,
        row.category,
        row.type,
        row.model,
        row.nodes,
        row.attachment
      ])
    ]);

    // Create workbook and append the worksheet
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '题目数据');

    // Generate Excel file
    XLSX.writeFile(workbook, `题目数据_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  return (
    <div className="min-h-screen bg-gray-50 font-sans text-gray-800">
      {/* 头部区域 */}
      <div className="bg-white shadow-sm">
        <div className="container mx-auto px-4 py-6">
          <h1 className="text-3xl font-bold text-center text-gray-800">题目生成器</h1>
        </div>
        <div className="border-b border-gray-200"></div>
      </div>

      {/* 主内容区 */}
      <div className="container mx-auto px-4 py-8">
        <div className="flex flex-col lg:flex-row gap-8">
          {/* 左侧配置区 */}
          <div className="flex-1 lg:w-[35%] bg-white rounded-lg shadow-sm p-6">
            {/* 模型选择 */}
            <div className="mb-6">
              <h2 className="text-lg font-semibold mb-4 flex items-center">
                <Cpu className="mr-2 text-blue-500" size={18} />
                模型选择
              </h2>
              <div className="flex gap-3">
                <button
                  onClick={() => setModel('DeepseekR1')}
                  className={`px-4 py-2 rounded-md flex-1 flex items-center justify-center ${model === 'DeepseekR1' ? 'bg-blue-500 text-white' : 'bg-gray-100 text-gray-700'}`}
                >
                  <FileCode2 className="mr-2" size={16} />
                  DeepseekR1
                </button>
                <button
                  onClick={() => setModel('GPT')}
                  className={`px-4 py-2 rounded-md flex-1 flex items-center justify-center ${model === 'GPT' ? 'bg-blue-500 text-white' : 'bg-gray-100 text-gray-700'}`}
                >
                  <FileTerminal className="mr-2" size={16} />
                  GPT
                </button>
              </div>
            </div>

            {/* 题目类型选择 */}
            <div className="mb-6">
              <h2 className="text-lg font-semibold mb-4 flex items-center">
                <BookOpen className="mr-2 text-blue-500" size={18} />
                题目类型
              </h2>
              
              {Object.entries({
                '基础知识': ['概念理解', '概念辨析'],
                '智力技能': ['数值计算', '数据分析', '文本处理', '指令遵循'],
                '认知策略': ['COT', '复杂推理', '编程'],
                '动作技能': ['RAG', 'SQL', 'API'],
                '情感感知': ['情绪', '态度', '观念', '安全']
              }).map(([category, types]) => (
                <div key={category} className="mb-3">
                  <button 
                    onClick={() => toggleCategory(category as QuestionCategory)}
                    className="w-full flex items-center justify-between p-3 bg-gray-50 rounded-md"
                  >
                    <div className="flex items-center">
                      {category === '基础知识' && <FileText className="mr-2 text-blue-500" size={16} />}
                      {category === '智力技能' && <Brain className="mr-2 text-blue-500" size={16} />}
                      {category === '认知策略' && <Activity className="mr-2 text-blue-500" size={16} />}
                      {category === '动作技能' && <FileSignature className="mr-2 text-blue-500" size={16} />}
                      {category === '情感感知' && <Heart className="mr-2 text-blue-500" size={16} />}
                      <span className="font-medium">{category}</span>
                    </div>
                    {expandedCategories[category] ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
                  </button>
                  
                  {expandedCategories[category] && (
                    <motion.div 
                      initial={{ opacity: 0, height: 0 }}
                      animate={{ opacity: 1, height: 'auto' }}
                      transition={{ duration: 0.2 }}
                      className="mt-2 grid grid-cols-1 gap-1 p-2"
                    >
                      {types.map(type => (
                        <button
                          key={type}
                          onClick={() => toggleQuestionType(category as QuestionCategory, type)}
                          className={`flex items-center p-2 rounded-md text-sm ${questionTypes[category as QuestionCategory]?.includes(type) ? 'bg-blue-50 text-blue-600' : 'bg-gray-50 text-gray-700'}`}
                        >
                          {questionTypes[category as QuestionCategory]?.includes(type) ? (
                            <Check className="mr-2" size={14} />
                          ) : (
                            <div className="w-4 h-4 border border-gray-300 rounded mr-2"></div>
                          )}
                          {type}
                        </button>
                      ))}
                    </motion.div>
                  )}
                </div>
              ))}
            </div>

            {/* 节点数量 */}
            <div className="mb-6">
              <h2 className="text-lg font-semibold mb-4 flex items-center">
                <FileStack className="mr-2 text-blue-500" size={18} />
                节点数量
              </h2>
              <div className="flex items-center gap-3">
                {[1, 2, 3, 4, 5].map(num => (
                  <button
                    key={num}
                    onClick={() => setNodeCount(num)}
                    className={`w-10 h-10 rounded-full flex items-center justify-center ${nodeCount === num ? 'bg-blue-500 text-white' : 'bg-gray-100 text-gray-700'}`}
                  >
                    {num}
                  </button>
                ))}
              </div>
            </div>

            {/* 是否需要附件 */}
            <div className="mb-6">
              <h2 className="text-lg font-semibold mb-4 flex items-center">
                <FileImage className="mr-2 text-blue-500" size={18} />
                是否需要附件
              </h2>
              <div className="flex gap-3">
                <button
                  onClick={() => setHasAttachment(true)}
                  className={`px-4 py-2 rounded-md flex-1 flex items-center justify-center ${hasAttachment ? 'bg-blue-500 text-white' : 'bg-gray-100 text-gray-700'}`}
                >
                  <Check className="mr-2" size={16} />
                  是
                </button>
                <button
                  onClick={() => setHasAttachment(false)}
                  className={`px-4 py-2 rounded-md flex-1 flex items-center justify-center ${!hasAttachment ? 'bg-blue-500 text-white' : 'bg-gray-100 text-gray-700'}`}
                >
                  <X className="mr-2" size={16} />
                  否
                </button>
              </div>
            </div>

            {/* 附件上传 */}
            {hasAttachment && (
              <div className="mb-6">
                <h2 className="text-lg font-semibold mb-4 flex items-center">
                  <FileUp className="mr-2 text-blue-500" size={18} />
                  PDF附件上传
                </h2>
                <label className="flex flex-col items-center justify-center w-full h-28 border-2 border-dashed border-gray-300 rounded-lg cursor-pointer bg-gray-50 hover:bg-gray-100">
                  <div className="flex flex-col items-center justify-center pt-4 pb-5">
                    <FileInput className="mb-1 text-gray-500" size={20} />
                    <p className="mb-1 text-sm text-gray-500">
                      <span className="font-semibold">点击上传</span> 或拖拽文件
                    </p>
                    <p className="text-xs text-gray-500">PDF文件 (最大10MB)</p>
                  </div>
                  <input 
                    type="file" 
                    className="hidden" 
                    accept=".pdf" 
                    onChange={handleFileChange}
                  />
                </label>
                {pdfFile && (
                  <div className="mt-2 flex items-center text-sm text-gray-600">
                    <File className="mr-2" size={14} />
                    {pdfFile.name}
                  </div>
                )}
              </div>
            )}

            {/* 生成题目数量 */}
            <div className="mb-6">
              <h2 className="text-lg font-semibold mb-4 flex items-center">
                <FileDigit className="mr-2 text-blue-500" size={18} />
                生成题目数量
              </h2>
              <div className="flex items-center gap-3">
                <button 
                  onClick={() => setQuestionCount(prev => Math.max(1, prev - 1))}
                  className="w-8 h-8 rounded-full bg-gray-100 flex items-center justify-center text-gray-700"
                >
                  <Minus size={14} />
                </button>
                <div className="text-lg font-medium">{questionCount}</div>
                <button 
                  onClick={() => setQuestionCount(prev => Math.min(10, prev + 1))}
                  className="w-8 h-8 rounded-full bg-gray-100 flex items-center justify-center text-gray-700"
                >
                  <Plus size={14} />
                </button>
              </div>
            </div>

            {/* 生成按钮 */}
            <button
              onClick={generateQuestions}
              disabled={isGenerating || Object.values(questionTypes).every(types => types?.length === 0)}
              className={`w-full py-3 rounded-lg flex items-center justify-center mt-4 ${isGenerating || Object.values(questionTypes).every(types => types?.length === 0) ? 'bg-gray-300 cursor-not-allowed' : 'bg-blue-500 hover:bg-blue-600 text-white'}`}
            >
              {isGenerating ? (
                <div className="flex items-center">
                  <svg className="animate-spin -ml-1 mr-3 h-4 w-4 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                  </svg>
                  生成中...
                </div>
              ) : (
                <>
                  <FilePlus className="mr-2" size={16} />
                  生成题目
                </>
              )}
            </button>
          </div>

          {/* 右侧预览区 */}
          <div className="flex-1 lg:w-[65%] bg-white rounded-lg shadow-sm p-6">
            <div className="flex items-center justify-between mb-4">
              <h2 className="text-lg font-semibold flex items-center">
                <FileSpreadsheet className="mr-2 text-blue-500" size={18} />
                Excel预览
              </h2>
              {generatedData.length > 0 && (
                <button 
                  onClick={downloadExcel}
                  className="flex items-center text-sm text-blue-500 hover:text-blue-600"
                >
                  <Download className="mr-1" size={14} />
                  下载Excel
                </button>
              )}
            </div>

            <div className="bg-gray-50 rounded-lg p-4 min-h-[400px]">
              {generatedData.length > 0 ? (
                <div className="overflow-x-auto">
                  <table className="min-w-full">
                    <thead>
                      <tr className="text-left text-sm text-gray-500 border-b border-gray-200">
                        <th className="pb-2">题目ID</th>
                        <th className="pb-2">分类</th>
                        <th className="pb-2">类型</th>
                        <th className="pb-2">模型</th>
                        <th className="pb-2">节点</th>
                        <th className="pb-2">附件</th>
                      </tr>
                    </thead>
                    <tbody>
                      {generatedData.map((row, index) => (
                        <tr key={index} className="text-sm border-b border-gray-100 hover:bg-gray-100">
                          <td className="py-2">{row.id}</td>
                          <td className="py-2">{row.category}</td>
                          <td className="py-2">{row.type}</td>
                          <td className="py-2">{row.model}</td>
                          <td className="py-2">{row.nodes}</td>
                          <td className="py-2">{row.attachment}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : (
                <div className="flex flex-col items-center justify-center h-full text-gray-400">
                  <FileSearch2 size={36} className="mb-3" />
                  <p>生成题目后将在此处预览</p>
                </div>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* 页脚区域 */}
      <div className="bg-gray-50 border-t border-gray-200">
        <div className="container mx-auto px-4 py-6 text-center text-sm text-gray-500">
          <p>created by <a href="https://space.coze.cn" className="text-blue-500 hover:text-blue-600">coze space</a> | 页面内容均由 AI 生成，仅供参考</p>
        </div>
      </div>
    </div>
  );
};

export default QuestionGenerator;