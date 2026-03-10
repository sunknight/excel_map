// 全局变量
let nodes = [];
let scale = 0.7; // 默认缩放级别设为0.7，让脑图显示更小
const MIN_SCALE = 0.3; // 最小缩放级别：30%
const MAX_SCALE = 2; // 最大缩放级别：200%
let panX = 0;
let panY = 0;
let isPanning = false;
let lastX, lastY;
let testData = [];
let mindmapData = null;
let nodeExpanded = new Set();
let workbook = null;
let currentSheet = null;
let sheetHeaders = []; // 存储当前sheet的表头
let fieldMapping = {
  level1: "",
  level2: "",
  level3: "",
  level4: "",
  number: "",
}; // 默认字段映射

// 默认字段名列表（支持多个可选字段名）
const DEFAULT_FIELD_NAMES = {
  level1: ["功能", "L1", "level1", "一级", "一级模块"],
  level2: ["类型", "L2", "level2", "二级", "二级模块"],
  level3: ["子功能", "L3", "level3", "三级", "三级模块", "用例类型"],
  level4: ["标题", "用例名称", "L4", "level4", "四级"],
  number: ["编号", "ID", "Number", "number"],
};

const viewport = document.getElementById("viewport");
const canvas = document.getElementById("canvas");
const svg = document.getElementById("connections");
const emptyState = document.getElementById("emptyState");

// Toast 提示函数
function showToast(message, type = "info") {
  const container = document.getElementById("toast-container");
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;

  const icons = {
    success: "✓",
    error: "✕",
    info: "ℹ",
  };

  toast.innerHTML = `<span class="toast-icon">${icons[type] || icons.info}</span><span>${message}</span>`;
  container.appendChild(toast);

  // 触发动画
  requestAnimationFrame(() => {
    toast.classList.add("show");
  });

  // 3秒后移除
  setTimeout(() => {
    toast.classList.remove("show");
    setTimeout(() => {
      if (container.contains(toast)) {
        container.removeChild(toast);
      }
    }, 300);
  }, 3000);
}

// Excel文件上传处理
document
  .getElementById("fileInput")
  .addEventListener("change", handleFileUpload);

function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  // 清空之前的脑图
  clearMindmap();

  document.getElementById("fileName").textContent = file.name;

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      workbook = XLSX.read(data, { type: "array" });

      // 获取所有工作表名称
      const sheetNames = workbook.SheetNames;

      if (sheetNames.length === 0) {
        alert("Excel文件中没有工作表");
        return;
      }

      // 填充Sheet选择器
      const sheetSelect = document.getElementById("sheetSelect");
      sheetSelect.innerHTML = '<option value="">请选择工作表</option>';
      sheetNames.forEach((name, index) => {
        const option = document.createElement("option");
        option.value = name;
        option.textContent = name;
        sheetSelect.appendChild(option);
      });

      // 显示Sheet选择器
      document.getElementById("sheetSelector").classList.add("active");

      // 不自动加载，让用户手动选择
      console.log(
        "已加载Excel文件，包含",
        sheetNames.length,
        "个工作表:",
        sheetNames.join(", "),
      );
    } catch (error) {
      console.error("解析Excel文件失败:", error);
      alert("解析Excel文件失败，请检查文件格式");
    }
  };
  reader.readAsArrayBuffer(file);
}

// 切换工作表
function switchSheet() {
  const sheetName = document.getElementById("sheetSelect").value;
  if (!sheetName || !workbook) return;

  currentSheet = sheetName;

  try {
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    if (jsonData.length < 2) {
      console.warn("工作表数据不足");
      // 清空脑图显示
      clearMindmap();
      return;
    }

    // 保存表头
    sheetHeaders = jsonData[0];

    // 自动填充字段选择器
    autoFillFieldSelectors();

    // 尝试自动匹配默认字段
    const autoMatched = autoMatchFields();

    if (autoMatched) {
      // 如果成功匹配必需字段，直接生成脑图
      console.log("已自动匹配字段并生成脑图");
      parseExcelData(jsonData);
    } else {
      // 如果未能匹配必需字段，清空之前的脑图并提示用户手动配置
      console.log("未能自动匹配所有必需字段，请手动配置");
      clearMindmapCanvas();
    }
  } catch (error) {
    console.error("解析工作表失败:", error);
  }
}

// 自动匹配字段
function autoMatchFields() {
  let matched = true;

  // 必需字段：level1 和 level4
  const requiredFields = ["level1", "level4"];

  // 匹配每个字段
  Object.keys(DEFAULT_FIELD_NAMES).forEach((fieldKey) => {
    const possibleNames = DEFAULT_FIELD_NAMES[fieldKey];
    const found = possibleNames.find((name) => sheetHeaders.includes(name));

    if (found) {
      fieldMapping[fieldKey] = found;
      // 更新下拉选择器
      const selectId =
        "field" + fieldKey.charAt(0).toUpperCase() + fieldKey.slice(1);
      const select = document.getElementById(selectId);
      if (select) {
        select.value = found;
      }
    } else if (requiredFields.includes(fieldKey)) {
      // 必需字段未找到
      matched = false;
    }
  });

  return matched;
}

// 自动填充字段选择器
function autoFillFieldSelectors() {
  const selectors = [
    "fieldLevel1",
    "fieldLevel2",
    "fieldLevel3",
    "fieldLevel4",
    "fieldNumber",
  ];

  selectors.forEach((selectorId) => {
    const select = document.getElementById(selectorId);
    select.innerHTML = '<option value="">请选择字段</option>';

    sheetHeaders.forEach((header) => {
      const option = document.createElement("option");
      option.value = header;
      option.textContent = header;
      select.appendChild(option);
    });
  });

  // 尝试自动匹配默认字段
  const defaultMappings = {
    fieldLevel1: fieldMapping.level1,
    fieldLevel2: fieldMapping.level2,
    fieldLevel3: fieldMapping.level3,
    fieldLevel4: fieldMapping.level4,
    fieldNumber: fieldMapping.number,
  };

  Object.keys(defaultMappings).forEach((selectorId) => {
    const select = document.getElementById(selectorId);
    const defaultValue = defaultMappings[selectorId];
    if (sheetHeaders.includes(defaultValue)) {
      select.value = defaultValue;
    }
  });
}

// 显示字段选择器弹窗
function showFieldSelector() {
  if (sheetHeaders.length === 0) {
    alert("请先选择工作表");
    return;
  }
  document.getElementById("fieldModal").style.display = "flex";
}

// 关闭字段选择器弹窗
function closeFieldModal() {
  document.getElementById("fieldModal").style.display = "none";
}

// 应用字段设置并生成脑图
function applyFieldSettings() {
  const level1Field = document.getElementById("fieldLevel1").value;
  const level2Field = document.getElementById("fieldLevel2").value;
  const level3Field = document.getElementById("fieldLevel3").value;
  const level4Field = document.getElementById("fieldLevel4").value;
  const numberField = document.getElementById("fieldNumber").value;

  // 验证至少选择一个字段
  if (!level1Field && !level2Field && !level3Field && !level4Field) {
    alert("请至少选择一个字段（Level 1-4）");
    return;
  }

  // 更新字段映射（未选择的字段为null）
  fieldMapping = {
    level1: level1Field || null,
    level2: level2Field || null,
    level3: level3Field || null,
    level4: level4Field || null,
    number: numberField || null,
  };

  // 关闭弹窗
  closeFieldModal();

  // 解析数据并生成脑图
  try {
    const worksheet = workbook.Sheets[currentSheet];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    parseExcelData(jsonData);
  } catch (error) {
    console.error("解析数据失败:", error);
    alert("解析数据失败: " + error.message);
  }
}

// 解析Excel数据
function parseExcelData(rawData) {
  const headers = rawData[0];
  const rows = rawData.slice(1);

  // 查找字段索引 - 使用用户选择的字段（未选择的为-1）
  const fieldMap = {
    number: fieldMapping.number ? headers.indexOf(fieldMapping.number) : -1,
    level1: fieldMapping.level1 ? headers.indexOf(fieldMapping.level1) : -1,
    level2: fieldMapping.level2 ? headers.indexOf(fieldMapping.level2) : -1,
    level3: fieldMapping.level3 ? headers.indexOf(fieldMapping.level3) : -1,
    level4: fieldMapping.level4 ? headers.indexOf(fieldMapping.level4) : -1,
  };

  // 提取数据 - 保存所有字段
  testData = rows
    .filter((row) => row.length > 0)
    .map((row, index) => {
      const testCase = {
        _level1: row[fieldMap.level1] || "未分类", // 功能
        _level2: row[fieldMap.level2] || "默认", // 类型
        _level3: row[fieldMap.level3] || "未分类", // 子功能
        _level4: row[fieldMap.level4] || "未命名", // 用例名称
        _number: row[fieldMap.number] || index + 2, // 编号
        _headers: headers, // 保存所有表头
        _rawData: row, // 保存原始数据
      };

      // 添加所有字段
      headers.forEach((header, idx) => {
        if (!testCase.hasOwnProperty(header) && row[idx] !== undefined) {
          testCase[header] = row[idx];
        }
      });

      return testCase;
    });

  // 生成脑图数据
  generateMindmapData();
  return true;
}

// 生成脑图数据结构
function generateMindmapData() {
  if (testData.length === 0) {
    alert("没有有效的测试用例数据");
    return;
  }

  // 按Level1-Level2-Level3分组，未选择的字段使用"默认"
  const functionGroups = {};

  testData.forEach((testCase) => {
    // 获取各级别的值，未选择字段的级别统一使用"默认"
    const level1 = fieldMapping.level1 ? testCase._level1 || "默认" : "默认";
    const level2 = fieldMapping.level2 ? testCase._level2 || "默认" : "默认";
    const level3 = fieldMapping.level3 ? testCase._level3 || "默认" : "默认";

    if (!functionGroups[level1]) {
      functionGroups[level1] = {};
    }

    if (!functionGroups[level1][level2]) {
      functionGroups[level1][level2] = {};
    }

    if (!functionGroups[level1][level2][level3]) {
      functionGroups[level1][level2][level3] = [];
    }

    functionGroups[level1][level2][level3].push(testCase);
  });

  // 构建树形结构
  mindmapData = {
    name: "测试用例",
    level: 0,
    children: [],
  };

  // 一级节点按A-Z排序
  Object.keys(functionGroups)
    .sort((a, b) => {
      return a.localeCompare(b, "zh-CN", { sensitivity: "base" });
    })
    .forEach((level1) => {
      const level1Node = {
        name: level1,
        level: 1,
        children: [],
      };

      // 二级节点按A-Z排序
      Object.keys(functionGroups[level1])
        .sort((a, b) => {
          return a.localeCompare(b, "zh-CN", { sensitivity: "base" });
        })
        .forEach((level2) => {
          const level2Node = {
            name: level2,
            level: 2,
            children: [],
          };

          // 三级节点按A-Z排序
          Object.keys(functionGroups[level1][level2])
            .sort((a, b) => {
              return a.localeCompare(b, "zh-CN", { sensitivity: "base" });
            })
            .forEach((level3) => {
              // 创建三级节点
              const level3Node = {
                name: level3,
                level: 3,
                children: [],
              };

              functionGroups[level1][level2][level3].forEach((testCase) => {
                const level4 = fieldMapping.level4
                  ? testCase._level4 || "默认"
                  : "默认";
                const number = testCase._number || "";

                level3Node.children.push({
                  name: `${number ? "#" + number + "：" : ""}${level4}`,
                  level: 4,
                  testCase: testCase,
                });
              });

              level2Node.children.push(level3Node);
            });

          level1Node.children.push(level2Node);
        });

      mindmapData.children.push(level1Node);
    });

  // 渲染脑图
  renderMindmap();
}

// 渲染脑图
function renderMindmap() {
  // 清空现有内容
  nodes = [];
  svg.innerHTML = "";
  canvas.innerHTML = '<svg class="connections" id="connections"></svg>';

  // 重置视图参数
  scale = 0.7;
  panX = 0;
  panY = 0;

  // 隐藏空状态
  emptyState.style.display = "none";

  // 启用控制按钮
  document.getElementById("resetBtn").disabled = false;
  document.getElementById("expandBtn").disabled = false;
  document.getElementById("collapseBtn").disabled = false;
  document.getElementById("autoLayoutBtn").disabled = false;
  document.getElementById("fitBtn").disabled = false;
  document.getElementById("copyBtn").disabled = false;
  document.getElementById("downloadBtn").disabled = false;

  // 创建节点
  createNodes(mindmapData, 0, -1);

  // 自动布局
  autoLayout();
}

// 创建节点 - 修复版本，确保正确的节点顺序
function createNodes(data, level, parentId) {
  // 先创建当前节点
  const nodeId = nodes.length;
  const node = {
    id: nodeId,
    name: data.name,
    level: level,
    x: 0,
    y: 0,
    width: 0,
    height: 0,
    parentId: parentId,
    children: [], // 先初始化为空数组
    testCase: data.testCase || null,
    expanded: true, // 默认展开所有节点
  };

  nodes.push(node);

  // 然后递归创建子节点，并记录子节点ID
  if (data.children && data.children.length > 0) {
    data.children.forEach((child) => {
      const childId = createNodes(child, level + 1, nodeId);
      node.children.push(childId);
    });
  }

  return nodeId;
}

// 自动布局 - 使用改进的树形布局算法，确保不重叠
function autoLayout(skipUpdateTransform = false) {
  if (!mindmapData || nodes.length === 0) return;

  const NODE_WIDTH = 220;
  const NODE_HEIGHT = 70;
  const HORIZONTAL_GAP = 150;
  const VERTICAL_GAP = 50;

  // 根据节点层级获取节点高度
  function getNodeHeight(node) {
    if (node.level === 4) {
      return 40; // 用例节点高度
    }
    return NODE_HEIGHT;
  }

  // 根据节点层级获取垂直间距
  function getVerticalGap(parentNode) {
    if (parentNode.level === 3) {
      return 15; // 类型到用例的间距更小
    }
    return VERTICAL_GAP;
  }

  // 第一步：计算每个节点的子树所需的总高度
  function calculateSubtreeSize(nodeId) {
    const node = nodes[nodeId];
    const nodeHeight = getNodeHeight(node);

    if (!node.expanded || node.children.length === 0) {
      // 叶子节点或折叠的节点
      node.subtreeHeight = nodeHeight;
      return nodeHeight;
    }

    // 展开的节点：计算所有子节点的高度总和
    let totalHeight = 0;
    const verticalGap = getVerticalGap(node);
    for (let i = 0; i < node.children.length; i++) {
      const childHeight = calculateSubtreeSize(node.children[i]);
      totalHeight += childHeight;
      if (i < node.children.length - 1) {
        totalHeight += verticalGap;
      }
    }

    node.subtreeHeight = Math.max(nodeHeight, totalHeight);
    return node.subtreeHeight;
  }

  // 第二步：设置节点位置
  function layoutNode(nodeId, depth, startY) {
    const node = nodes[nodeId];
    const nodeHeight = getNodeHeight(node);

    // 设置X坐标（基于深度）
    node.x = 100 + depth * (NODE_WIDTH + HORIZONTAL_GAP);

    // 设置Y坐标
    if (!node.expanded || node.children.length === 0) {
      // 叶子节点：直接使用startY
      node.y = startY;
    } else {
      // 有子节点的节点：先布局子节点，然后根据子节点的实际位置计算父节点位置
      let currentY = startY;
      const verticalGap = getVerticalGap(node);

      // 先递归布局所有子节点
      for (let i = 0; i < node.children.length; i++) {
        const childId = node.children[i];
        const child = nodes[childId];
        const childHeight = child.subtreeHeight;

        // 布局子节点
        layoutNode(childId, depth + 1, currentY);

        currentY += childHeight + verticalGap;
      }

      // 现在所有子节点都已经布局完成，可以获取它们的实际位置
      const firstChildId = node.children[0];
      const firstChild = nodes[firstChildId];
      const firstChildNodeHeight = getNodeHeight(firstChild);
      const firstChildCenter = firstChild.y + firstChildNodeHeight / 2;

      const lastChildId = node.children[node.children.length - 1];
      const lastChild = nodes[lastChildId];
      const lastChildNodeHeight = getNodeHeight(lastChild);
      const lastChildCenter = lastChild.y + lastChildNodeHeight / 2;

      // 父节点中心 = 第一个子节点中心 + (最后一个子节点中心 - 第一个子节点中心) / 2
      node.y = (firstChildCenter + lastChildCenter) / 2 - nodeHeight / 2;
    }
  }

  // 执行布局
  const totalHeight = calculateSubtreeSize(0);

  // 计算起始Y坐标，使整个树垂直居中
  const viewportHeight = viewport.clientHeight;
  const startY = Math.max(50, (viewportHeight - totalHeight) / 2);

  layoutNode(0, 0, startY);

  // 渲染节点
  renderNodes();
  renderConnections();

  // 更新画布变换（除非跳过）
  if (!skipUpdateTransform) {
    updateCanvasTransform();
  }
}

// 计算节点的子用例总数（只计算level 4的测试用例节点）
function countChildTestCases(nodeId) {
  const node = nodes[nodeId];

  // 如果是level 4（测试用例节点），返回1
  if (node.level === 4) {
    return 1;
  }

  // 如果没有子节点，返回0
  if (!node.children || node.children.length === 0) {
    return 0;
  }

  // 递归计算所有子节点的用例数量
  let count = 0;
  node.children.forEach((childId) => {
    count += countChildTestCases(childId);
  });

  return count;
}

// 渲染节点
function renderNodes() {
  // 保留SVG连接线容器
  const connectionsSvg = canvas.querySelector("#connections");
  canvas.innerHTML = "";
  canvas.appendChild(connectionsSvg);

  // 检查节点的所有祖先节点是否都展开
  function shouldRenderNode(node) {
    if (node.parentId === -1) return true; // 根节点总是渲染

    let currentNode = node;
    while (currentNode.parentId !== -1) {
      const parentNode = nodes[currentNode.parentId];
      if (!parentNode.expanded) {
        return false; // 如果任何一个祖先节点折叠，不渲染
      }
      currentNode = parentNode;
    }
    return true;
  }

  nodes.forEach((node) => {
    if (!shouldRenderNode(node)) {
      return; // 如果祖先节点有折叠的，不渲染
    }

    const nodeElement = document.createElement("div");
    nodeElement.className = `node level-${Math.min(node.level, 4)}`;
    nodeElement.id = `node-${node.id}`;

    // 添加展开/折叠指示器（如果有子节点）
    let indicator = "";
    if (node.children && node.children.length > 0) {
      indicator = node.expanded ? "▼ " : "▶ ";
    }

    // 为level 1, 2, 3添加子用例数量统计
    let displayText = node.name;
    if (
      node.level >= 1 &&
      node.level <= 3 &&
      node.children &&
      node.children.length > 0
    ) {
      const testCaseCount = countChildTestCases(node.id);
      displayText = `${node.name} (${testCaseCount})`;
    }

    nodeElement.innerHTML = `<div class="node-text">${indicator}${displayText}</div>`;

    // 设置位置
    nodeElement.style.left = `${node.x}px`;
    nodeElement.style.top = `${node.y}px`;

    // 设置鼠标样式
    if (node.children && node.children.length > 0) {
      nodeElement.style.cursor = "pointer";
    }

    // 添加拖拽事件（包含点击处理）
    addDragEvents(nodeElement, node.id);

    canvas.appendChild(nodeElement);
  });
}

// 渲染连接线
function renderConnections() {
  const svg = document.getElementById("connections");
  svg.innerHTML = "";

  nodes.forEach((node) => {
    if (node.parentId === -1) return; // 根节点没有父节点

    const parentNode = nodes[node.parentId];
    if (!parentNode.expanded) return; // 如果父节点折叠，不渲染连接线

    const startX = parentNode.x + getNodeWidth(parentNode);
    const startY = parentNode.y + getNodeHeight(parentNode) / 2;
    const endX = node.x;
    const endY = node.y + getNodeHeight(node) / 2;

    // 创建贝塞尔曲线
    const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    const controlX1 = startX + (endX - startX) / 2;
    const controlY1 = startY;
    const controlX2 = startX + (endX - startX) / 2;
    const controlY2 = endY;

    const d = `M ${startX} ${startY} C ${controlX1} ${controlY1}, ${controlX2} ${controlY2}, ${endX} ${endY}`;

    path.setAttribute("d", d);
    path.setAttribute("class", "connection");
    svg.appendChild(path);
  });
}

// 获取节点宽度
function getNodeWidth(node) {
  const nodeElement = document.getElementById(`node-${node.id}`);
  if (nodeElement) {
    return nodeElement.offsetWidth;
  }
  return 220; // 默认宽度
}

// 获取节点高度
function getNodeHeight(node) {
  const nodeElement = document.getElementById(`node-${node.id}`);
  if (nodeElement) {
    return nodeElement.offsetHeight;
  }
  return 60; // 默认高度
}

// 添加拖拽事件
function addDragEvents(element, nodeId) {
  let isDragging = false;
  let hasMoved = false;
  let startX, startY, initialX, initialY;

  element.addEventListener("mousedown", (e) => {
    if (e.button !== 0) return; // 只响应左键
    isDragging = true;
    hasMoved = false;
    element.classList.add("dragging");
    startX = e.clientX;
    startY = e.clientY;
    initialX = nodes[nodeId].x;
    initialY = nodes[nodeId].y;
    e.preventDefault();
  });

  document.addEventListener("mousemove", (e) => {
    if (!isDragging) return;

    const dx = e.clientX - startX;
    const dy = e.clientY - startY;

    // 如果移动距离超过5px，认为是拖拽而不是点击
    if (Math.abs(dx) > 5 || Math.abs(dy) > 5) {
      hasMoved = true;
    }

    nodes[nodeId].x = initialX + dx;
    nodes[nodeId].y = initialY + dy;

    element.style.left = `${nodes[nodeId].x}px`;
    element.style.top = `${nodes[nodeId].y}px`;

    renderConnections();
  });

  document.addEventListener("mouseup", (e) => {
    if (isDragging) {
      isDragging = false;
      element.classList.remove("dragging");

      // 如果没有移动，触发点击事件
      if (!hasMoved) {
        const node = nodes[nodeId];
        if (node.testCase) {
          showTestCaseDetail(node.testCase);
        } else if (node.children && node.children.length > 0) {
          toggleNode(nodeId);
        }
      }
    }
  });
}

// 显示测试用例详情
function showTestCaseDetail(testCase) {
  const modal = document.getElementById("modal");
  const modalTitle = document.getElementById("modal-title");
  const modalContent = document.getElementById("modal-content");

  // 使用用例名称作为标题
  modalTitle.textContent = testCase._level4 || "测试用例详情";

  // 获取所有字段（排除内部字段）
  const displayFields = Object.keys(testCase).filter(
    (key) =>
      !key.startsWith("_") &&
      testCase[key] !== undefined &&
      testCase[key] !== "",
  );

  // 构建内容
  let content =
    '<div style="display: grid; grid-template-columns: 120px 1fr; gap: 12px 20px; font-size: 13px;">';

  displayFields.forEach((key) => {
    const value = testCase[key];
    const displayValue = value || "未设置";

    // 特殊处理优先级字段
    if (key === "优先级") {
      const priorityClass = value ? `priority-${value}` : "";
      content += `
                <div style="font-weight: 600; color: #262626; text-align: right; padding-top: 4px;">${key}</div>
                <div style="color: #595959; line-height: 1.6;"><span class="priority-tag ${priorityClass}">${displayValue}</span></div>
            `;
    } else {
      // 判断是否需要保留换行
      const needsPreWrap = [
        "业务流程",
        "前置条件",
        "步骤描述",
        "预期结果",
        "备注",
      ].includes(key);
      const style = needsPreWrap
        ? "white-space: pre-wrap; word-break: break-word;"
        : "";

      content += `
                <div style="font-weight: 600; color: #262626; text-align: right; padding-top: 4px;">${key}</div>
                <div style="color: #595959; line-height: 1.6; ${style}">${displayValue}</div>
            `;
    }
  });

  content += "</div>";

  modalContent.innerHTML = content;
  modal.style.display = "flex";
}

// 关闭弹窗
function closeModal() {
  document.getElementById("modal").style.display = "none";
}

// 清空脑图
function clearMindmap() {
  // 清空所有数据
  nodes = [];
  testData = [];
  mindmapData = null;
  nodeExpanded = {};
  workbook = null;
  currentSheet = null;
  sheetHeaders = [];

  // 重置视图
  scale = 1;
  panX = 0;
  panY = 0;

  // 清空SVG内容
  svg.innerHTML = "";
  svg.style.display = "none";
  emptyState.style.display = "flex";

  // 重置文件名显示
  document.getElementById("fileName").textContent = "";

  // 清空工作表选择器
  const sheetSelect = document.getElementById("sheetSelect");
  if (sheetSelect) {
    sheetSelect.innerHTML = "";
  }

  // 禁用控制按钮
  document.getElementById("resetBtn").disabled = true;
  document.getElementById("expandBtn").disabled = true;
  document.getElementById("collapseBtn").disabled = true;
  document.getElementById("autoLayoutBtn").disabled = true;
  document.getElementById("fitBtn").disabled = true;
  document.getElementById("copyBtn").disabled = true;
  document.getElementById("downloadBtn").disabled = true;
}

// 清空脑图画布（保留文件和 workbook）
function clearMindmapCanvas() {
  // 清空脑图数据
  nodes = [];
  testData = [];
  mindmapData = null;
  nodeExpanded = {};

  // 重置视图
  scale = 0.7;
  panX = 0;
  panY = 0;

  // 清空画布内容
  canvas.innerHTML = '<svg class="connections" id="connections"></svg>';
  svg.innerHTML = "";
  svg.style.display = "none";

  // 显示空状态
  emptyState.style.display = "flex";

  // 禁用控制按钮
  document.getElementById("resetBtn").disabled = true;
  document.getElementById("expandBtn").disabled = true;
  document.getElementById("collapseBtn").disabled = true;
  document.getElementById("autoLayoutBtn").disabled = true;
  document.getElementById("fitBtn").disabled = true;
  document.getElementById("copyBtn").disabled = true;
  document.getElementById("downloadBtn").disabled = true;
}

// 重置布局
function resetLayout() {
  if (!mindmapData) return;
  scale = 0.7; // 重置时也使用0.7的缩放级别
  panX = 0;
  panY = 0;
  autoLayout();
}

// 切换节点展开/折叠状态
function toggleNode(nodeId) {
  const node = nodes[nodeId];
  if (!node || !node.children || node.children.length === 0) return;

  // 保存当前节点在屏幕上的位置（在 DOM 更新之前）
  const nodeElement = document.getElementById(`node-${nodeId}`);
  const oldRect = nodeElement.getBoundingClientRect();
  const oldCenterX = oldRect.left + oldRect.width / 2;
  const oldCenterY = oldRect.top + oldRect.height / 2;

  if (node.expanded) {
    // 折叠：折叠当前节点及其所有子孙节点
    node.expanded = false;
    collapseAllDescendants(nodeId);
  } else {
    // 展开：只展开当前节点，不展开子节点
    node.expanded = true;
  }

  // 重新布局，但不更新画布变换
  autoLayout(true);

  // 使用 requestAnimationFrame 确保 DOM 完全渲染后再获取新位置
  requestAnimationFrame(() => {
    const newElement = document.getElementById(`node-${nodeId}`);
    const newRect = newElement.getBoundingClientRect();
    const newCenterX = newRect.left + newRect.width / 2;
    const newCenterY = newRect.top + newRect.height / 2;

    // 计算节点在屏幕上移动的距离
    const deltaX = newCenterX - oldCenterX;
    const deltaY = newCenterY - oldCenterY;

    // 调整 panX 和 panY 来补偿移动
    panX -= deltaX;
    panY -= deltaY;

    // 更新画布变换
    updateCanvasTransform();
  });
}

// 递归折叠所有子孙节点
function collapseAllDescendants(nodeId) {
  const node = nodes[nodeId];
  if (!node || !node.children) return;

  node.children.forEach((childId) => {
    const child = nodes[childId];
    if (child) {
      child.expanded = false;
      collapseAllDescendants(childId);
    }
  });
}

// 展开所有节点
function expandAll() {
  nodes.forEach((node) => {
    node.expanded = true;
  });
  autoLayout();
}

// 折叠所有节点 - 只折叠非根节点
function collapseAll() {
  nodes.forEach((node) => {
    if (node.level > 0) {
      node.expanded = false;
    } else {
      node.expanded = true; // 根节点保持展开
    }
  });
  autoLayout();
}

// 适应屏幕
function fitToScreen() {
  if (nodes.length === 0) return;

  // 计算边界
  let minX = Infinity,
    minY = Infinity,
    maxX = -Infinity,
    maxY = -Infinity;

  nodes.forEach((node) => {
    minX = Math.min(minX, node.x);
    minY = Math.min(minY, node.y);
    maxX = Math.max(maxX, node.x + getNodeWidth(node));
    maxY = Math.max(maxY, node.y + getNodeHeight(node));
  });

  const contentWidth = maxX - minX;
  const contentHeight = maxY - minY;
  const viewportWidth = viewport.clientWidth;
  const viewportHeight = viewport.clientHeight;

  const scaleX = (viewportWidth - 100) / contentWidth;
  const scaleY = (viewportHeight - 100) / contentHeight;
  scale = Math.min(scaleX, scaleY, 1);

  panX = (viewportWidth - contentWidth * scale) / 2 - minX * scale;
  panY = (viewportHeight - contentHeight * scale) / 2 - minY * scale;

  updateCanvasTransform();
}

// 更新画布变换
function updateCanvasTransform() {
  canvas.style.transform = `translate(${panX}px, ${panY}px) scale(${scale})`;
}

// 缩放控制
function zoomIn() {
  scale = Math.min(scale * 1.2, MAX_SCALE);
  updateCanvasTransform();
}

function zoomOut() {
  scale = Math.max(scale / 1.2, MIN_SCALE);
  updateCanvasTransform();
}

function resetZoom() {
  scale = 1;
  panX = 0;
  panY = 0;
  updateCanvasTransform();
}

// 画布拖拽平移
viewport.addEventListener("mousedown", (e) => {
  if (
    e.target === viewport ||
    e.target === canvas ||
    e.target.tagName === "svg"
  ) {
    isPanning = true;
    lastX = e.clientX;
    lastY = e.clientY;
    viewport.style.cursor = "grabbing";
  }
});

document.addEventListener("mousemove", (e) => {
  if (isPanning) {
    const dx = e.clientX - lastX;
    const dy = e.clientY - lastY;
    panX += dx;
    panY += dy;
    lastX = e.clientX;
    lastY = e.clientY;
    updateCanvasTransform();
  }
});

document.addEventListener("mouseup", () => {
  if (isPanning) {
    isPanning = false;
    viewport.style.cursor = "default";
  }
});

// 鼠标滚轮/触摸板 - 支持平移和缩放
viewport.addEventListener(
  "wheel",
  (e) => {
    e.preventDefault();

    // 检测是否是触控板缩放手势（ctrlKey 为 true 表示缩放）
    if (e.ctrlKey || e.metaKey) {
      // 触控板两指缩放
      // deltaY > 0 表示缩小，deltaY < 0 表示放大
      const zoomFactor = e.deltaY > 0 ? 0.9 : 1.1;

      // 计算鼠标位置相对于画布的坐标
      const rect = viewport.getBoundingClientRect();
      const mouseX = e.clientX - rect.left;
      const mouseY = e.clientY - rect.top;

      // 转换为画布坐标系
      const canvasX = (mouseX - panX) / scale;
      const canvasY = (mouseY - panY) / scale;

      // 应用缩放
      const newScale = Math.max(
        MIN_SCALE,
        Math.min(MAX_SCALE, scale * zoomFactor),
      );

      // 调整 panX 和 panY，使缩放以鼠标位置为中心
      panX = mouseX - canvasX * newScale;
      panY = mouseY - canvasY * newScale;
      scale = newScale;
    } else {
      // 普通滚轮或触控板双指滑动 - 移动画布
      panX -= e.deltaX;
      panY -= e.deltaY;
    }

    updateCanvasTransform();
  },
  { passive: false },
);

// 点击弹窗外部关闭
document.getElementById("modal").addEventListener("click", (e) => {
  if (e.target.id === "modal") {
    closeModal();
  }
});

// ESC键关闭弹窗
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    closeModal();
  }
});

// 复制画布到剪贴板
async function copyCanvas() {
  try {
    const canvas = document.getElementById("canvas");
    const svg = document.getElementById("connections");

    // 使用html2canvas库将DOM转换为canvas
    // 如果没有html2canvas，使用原生方法
    if (typeof html2canvas === "undefined") {
      // 动态加载html2canvas库
      await loadScript(
        "https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js",
      );
    }

    // 计算所有节点的边界范围
    const bounds = calculateCanvasBounds();

    // 临时调整画布样式以显示完整内容
    const originalTransform = canvas.style.transform;
    const originalWidth = canvas.style.width;
    const originalHeight = canvas.style.height;
    const originalOverflow = canvas.style.overflow;
    const originalPosition = canvas.style.position;
    const originalLeft = canvas.style.left;
    const originalTop = canvas.style.top;

    // 保存SVG原始样式
    const originalSvgWidth = svg.style.width;
    const originalSvgHeight = svg.style.height;
    const originalSvgViewBox = svg.getAttribute("viewBox");

    // 将画布移到屏幕外，避免样式调整时的视觉跳动
    canvas.style.position = "absolute";
    canvas.style.left = "-99999px";
    canvas.style.top = "-99999px";

    // 设置画布为完整内容大小
    canvas.style.transform = "none";
    canvas.style.width = `${bounds.width + 100}px`;
    canvas.style.height = `${bounds.height + 100}px`;
    canvas.style.overflow = "visible";

    // 设置SVG的宽高和viewBox以匹配画布大小
    svg.style.width = `${bounds.width + 100}px`;
    svg.style.height = `${bounds.height + 100}px`;
    svg.setAttribute(
      "viewBox",
      `0 0 ${bounds.width + 100} ${bounds.height + 100}`,
    );

    // 等待DOM更新
    await new Promise((resolve) => setTimeout(resolve, 100));

    // 使用html2canvas生成图片
    const canvasElement = await html2canvas(canvas, {
      backgroundColor: "#f5f7fa",
      scale: 2, // 提高清晰度
      useCORS: true,
      logging: false,
      width: bounds.width + 100,
      height: bounds.height + 100,
      x: 0,
      y: 0,
      scrollX: 0,
      scrollY: 0,
      allowTaint: true,
      foreignObjectRendering: false,
    });

    // 恢复画布样式
    canvas.style.transform = originalTransform;
    canvas.style.width = originalWidth;
    canvas.style.height = originalHeight;
    canvas.style.overflow = originalOverflow;
    canvas.style.position = originalPosition;
    canvas.style.left = originalLeft;
    canvas.style.top = originalTop;

    // 恢复SVG样式
    svg.style.width = originalSvgWidth;
    svg.style.height = originalSvgHeight;
    if (originalSvgViewBox) {
      svg.setAttribute("viewBox", originalSvgViewBox);
    } else {
      svg.removeAttribute("viewBox");
    }

    // 转换为blob
    canvasElement.toBlob(async (blob) => {
      try {
        // 复制到剪贴板
        await navigator.clipboard.write([
          new ClipboardItem({ "image/png": blob }),
        ]);
        showToast("画布已复制到剪贴板", "success");
      } catch (err) {
        console.error("复制失败:", err);
        showToast("复制失败，请使用下载功能", "error");
      }
    }, "image/png");
  } catch (error) {
    console.error("复制画布失败:", error);
    alert("复制失败: " + error.message);
  }
}

// 下载画布为PNG图片
async function downloadPNG() {
  try {
    const canvas = document.getElementById("canvas");
    const svg = document.getElementById("connections");

    // 使用html2canvas库将DOM转换为canvas
    if (typeof html2canvas === "undefined") {
      // 动态加载html2canvas库
      await loadScript(
        "https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js",
      );
    }

    // 计算所有节点的边界范围
    const bounds = calculateCanvasBounds();

    // 临时调整画布样式以显示完整内容
    const originalTransform = canvas.style.transform;
    const originalWidth = canvas.style.width;
    const originalHeight = canvas.style.height;
    const originalOverflow = canvas.style.overflow;
    const originalPosition = canvas.style.position;
    const originalLeft = canvas.style.left;
    const originalTop = canvas.style.top;

    // 保存SVG原始样式
    const originalSvgWidth = svg.style.width;
    const originalSvgHeight = svg.style.height;
    const originalSvgViewBox = svg.getAttribute("viewBox");

    // 将画布移到屏幕外，避免样式调整时的视觉跳动
    canvas.style.position = "absolute";
    canvas.style.left = "-99999px";
    canvas.style.top = "-99999px";

    // 设置画布为完整内容大小
    canvas.style.transform = "none";
    canvas.style.width = `${bounds.width + 100}px`;
    canvas.style.height = `${bounds.height + 100}px`;
    canvas.style.overflow = "visible";

    // 设置SVG的宽高和viewBox以匹配画布大小
    svg.style.width = `${bounds.width + 100}px`;
    svg.style.height = `${bounds.height + 100}px`;
    svg.setAttribute(
      "viewBox",
      `0 0 ${bounds.width + 100} ${bounds.height + 100}`,
    );

    // 等待DOM更新
    await new Promise((resolve) => setTimeout(resolve, 100));

    // 使用html2canvas生成图片
    const canvasElement = await html2canvas(canvas, {
      backgroundColor: "#f5f7fa",
      scale: 2, // 提高清晰度
      useCORS: true,
      logging: false,
      width: bounds.width + 100,
      height: bounds.height + 100,
      x: 0,
      y: 0,
      scrollX: 0,
      scrollY: 0,
      allowTaint: true,
      foreignObjectRendering: false,
    });

    // 恢复画布样式
    canvas.style.transform = originalTransform;
    canvas.style.width = originalWidth;
    canvas.style.height = originalHeight;
    canvas.style.overflow = originalOverflow;
    canvas.style.position = originalPosition;
    canvas.style.left = originalLeft;
    canvas.style.top = originalTop;

    // 恢复SVG样式
    svg.style.width = originalSvgWidth;
    svg.style.height = originalSvgHeight;
    if (originalSvgViewBox) {
      svg.setAttribute("viewBox", originalSvgViewBox);
    } else {
      svg.removeAttribute("viewBox");
    }

    // 转换为图片并下载
    canvasElement.toBlob((blob) => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `测试用例脑图_${new Date().toLocaleDateString("zh-CN").replace(/\//g, "-")}.png`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, "image/png");
  } catch (error) {
    console.error("下载PNG失败:", error);
    alert("下载失败: " + error.message);
  }
}

// 计算画布边界范围
function calculateCanvasBounds() {
  if (nodes.length === 0) {
    return { width: 800, height: 600 };
  }

  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;

  nodes.forEach((node) => {
    const nodeElement = document.getElementById(`node-${node.id}`);
    if (nodeElement) {
      const rect = nodeElement.getBoundingClientRect();
      const canvasRect = canvas.getBoundingClientRect();

      // 计算节点相对于画布的位置
      const nodeX = node.x;
      const nodeY = node.y;
      const nodeWidth = nodeElement.offsetWidth;
      const nodeHeight = nodeElement.offsetHeight;

      minX = Math.min(minX, nodeX);
      minY = Math.min(minY, nodeY);
      maxX = Math.max(maxX, nodeX + nodeWidth);
      maxY = Math.max(maxY, nodeY + nodeHeight);
    }
  });

  // 添加边距
  const padding = 50;
  return {
    width: maxX - minX + padding * 2,
    height: maxY - minY + padding * 2,
    offsetX: minX - padding,
    offsetY: minY - padding,
  };
}

// 动态加载脚本
function loadScript(src) {
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = src;
    script.onload = resolve;
    script.onerror = reject;
    document.head.appendChild(script);
  });
}
