/* eslint-disable @typescript-eslint/no-unused-vars */
/// <reference types="office-js" />

import './taskpane.css';

// Shape 정보 인터페이스
interface ShapeInfo {
  id: string;
  name: string;
  type: string;
  index: number;
}

// 현재 슬라이드의 Shape 목록
let currentShapes: ShapeInfo[] = [];

// Office 초기화
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    initializeUI();
  }
});

/**
 * UI 초기화
 */
function initializeUI(): void {
  document.getElementById('analyzeButton')?.addEventListener('click', analyzeSlide);
  document.getElementById('setAltTextButton')?.addEventListener('click', setAltTextForSelected);
  document.getElementById('applyRulesButton')?.addEventListener('click', () => applyRulesToAll(false));
  document.getElementById('applyRulesEmptyButton')?.addEventListener('click', () => applyRulesToAll(true));
  document.getElementById('ruleType')?.addEventListener('change', updateRuleUI);
}

/**
 * 상태 메시지 표시
 */
function showStatus(message: string, type: 'info' | 'success' | 'error' = 'info'): void {
  const statusDiv = document.getElementById('statusMessage');
  if (!statusDiv) return;

  statusDiv.className = `status-message ${type}`;
  statusDiv.textContent = message;

  setTimeout(() => {
    statusDiv.textContent = '';
    statusDiv.className = '';
  }, 5000);
}

/**
 * 규칙 UI 업데이트
 */
function updateRuleUI(): void {
  const ruleType = (document.getElementById('ruleType') as HTMLSelectElement)?.value;
  const prefixGroup = document.getElementById('prefixGroup');
  const customGroup = document.getElementById('customTemplateGroup');

  if (prefixGroup) {
    prefixGroup.style.display = ruleType === 'prefix' ? 'block' : 'none';
  }
  if (customGroup) {
    customGroup.style.display = ruleType === 'custom' ? 'block' : 'none';
  }
}

/**
 * 슬라이드 분석
 */
async function analyzeSlide(): Promise<void> {
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      const slide = slides.getItemAt(0);
      const shapes = slide.shapes;

      shapes.load('items');
      await context.sync();

      currentShapes = [];
      const shapesList = document.getElementById('shapesList');
      if (!shapesList) return;

      shapesList.innerHTML = '';

      if (shapes.items.length === 0) {
        shapesList.innerHTML = '<div class="status-message info">이 슬라이드에는 요소가 없습니다.</div>';
        return;
      }

      for (let i = 0; i < shapes.items.length; i++) {
        const shape = shapes.items[i];
        shape.load('name,type,id');
        await context.sync();

        const shapeInfo: ShapeInfo = {
          id: shape.id,
          name: shape.name,
          type: shape.type,
          index: i
        };

        currentShapes.push(shapeInfo);

        const shapeDiv = document.createElement('div');
        shapeDiv.className = 'shape-item';
        shapeDiv.innerHTML = `
          <div class="shape-name">${shapeInfo.name}</div>
          <div class="shape-type">타입: ${getShapeTypeName(shapeInfo.type)} | 인덱스: ${i}</div>
        `;

        shapeDiv.onclick = () => selectShape(i);
        shapesList.appendChild(shapeDiv);
      }

      showStatus(`${shapes.items.length}개의 요소를 찾았습니다.`, 'success');
    });
  } catch (error) {
    showStatus('슬라이드 분석 중 오류 발생: ' + (error as Error).message, 'error');
    console.error(error);
  }
}

/**
 * Shape 타입 이름 변환
 */
function getShapeTypeName(type: string): string {
  const typeMap: Record<string, string> = {
    'GeometricShape': '도형',
    'Image': '이미지',
    'Group': '그룹',
    'Line': '선',
    'Table': '표',
    'Chart': '차트',
    'TextBox': '텍스트 상자',
    'ContentPlaceholder': '콘텐츠 자리 표시자',
    'Picture': '그림',
    'Unsupported': '지원되지 않는 타입'
  };
  return typeMap[type] || type;
}

/**
 * Shape 선택
 */
async function selectShape(index: number): Promise<void> {
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      const slide = slides.getItemAt(0);
      const shapes = slide.shapes;
      const shape = shapes.getItemAt(index);

      shape.load('name');
      await context.sync();

      showStatus(`선택됨: ${shape.name}`, 'info');
    });
  } catch (error) {
    showStatus('요소 선택 중 오류 발생: ' + (error as Error).message, 'error');
    console.error(error);
  }
}

/**
 * 선택한 요소에 대체 텍스트 설정
 */
async function setAltTextForSelected(): Promise<void> {
  const altText = (document.getElementById('altText') as HTMLTextAreaElement)?.value.trim();

  if (!altText) {
    showStatus('대체 텍스트를 입력해주세요.', 'error');
    return;
  }

  try {
    await PowerPoint.run(async (context) => {
      const selectedShapes = context.presentation.getSelectedShapes();
      const shapeCount = selectedShapes.getCount();

      await context.sync();

      if (shapeCount.value === 0) {
        showStatus('요소를 선택해주세요.', 'error');
        return;
      }

      selectedShapes.load('items');
      await context.sync();

      for (const shape of selectedShapes.items) {
        const currentName = shape.name;
        const baseName = currentName.replace(/\[ALT:.*?\]\s*/, '');
        shape.name = `[ALT: ${altText}] ${baseName}`;
      }

      await context.sync();

      showStatus(`${shapeCount.value}개 요소에 대체 텍스트를 설정했습니다.`, 'success');
      await analyzeSlide();
    });
  } catch (error) {
    showStatus('대체 텍스트 설정 중 오류 발생: ' + (error as Error).message, 'error');
    console.error(error);
  }
}

/**
 * 모든 요소에 규칙 적용
 */
async function applyRulesToAll(onlyEmpty: boolean): Promise<void> {
  const ruleType = (document.getElementById('ruleType') as HTMLSelectElement)?.value;

  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      const slide = slides.getItemAt(0);
      const shapes = slide.shapes;

      shapes.load('items');
      await context.sync();

      let count = 0;

      for (let i = 0; i < shapes.items.length; i++) {
        const shape = shapes.items[i];
        shape.load('name,type');
        await context.sync();

        const currentName = shape.name;
        const hasAlt = currentName.includes('[ALT:');

        if (onlyEmpty && hasAlt) {
          continue;
        }

        const baseName = currentName.replace(/\[ALT:.*?\]\s*/, '');
        let altText = '';

        switch (ruleType) {
          case 'prefix':
            const prefix = (document.getElementById('prefix') as HTMLInputElement)?.value || '';
            altText = prefix + baseName;
            break;

          case 'type':
            const typeName = getShapeTypeName(shape.type);
            altText = `${typeName}: ${baseName}`;
            break;

          case 'custom':
            const template = (document.getElementById('customTemplate') as HTMLInputElement)?.value || '';
            altText = template
              .replace('{name}', baseName)
              .replace('{type}', getShapeTypeName(shape.type))
              .replace('{index}', (i + 1).toString());
            break;
        }

        shape.name = `[ALT: ${altText}] ${baseName}`;
        count++;
      }

      await context.sync();

      showStatus(`${count}개 요소에 규칙을 적용했습니다.`, 'success');
      await analyzeSlide();
    });
  } catch (error) {
    showStatus('규칙 적용 중 오류 발생: ' + (error as Error).message, 'error');
    console.error(error);
  }
}
