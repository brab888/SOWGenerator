/** @jsxImportSource @emotion/react */
import React, { useState, useRef, useEffect, useCallback } from 'react';
import { useEditor, EditorContent } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import TextStyle from '@tiptap/extension-text-style';
import TextAlign from '@tiptap/extension-text-align';
import Color from '@tiptap/extension-color';
import * as XLSX from 'xlsx';
import styled from '@emotion/styled';
import {
  DndContext,
  closestCenter,
  KeyboardSensor,
  PointerSensor,
  useSensor,
  useSensors,
  DragEndEvent,
} from '@dnd-kit/core';
import {
  arrayMove,
  SortableContext,
  sortableKeyboardCoordinates,
  verticalListSortingStrategy,
  useSortable,
} from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';
import logo from './logo.svg';
import './App.css';
import { CellMappingPanel } from './components/CellMappingPanel';

interface Theme {
  '--bg-primary': string;
  '--bg-secondary': string;
  '--bg-tertiary': string;
  '--text-primary': string;
  '--text-secondary': string;
  '--text-tertiary': string;
  '--accent-primary': string;
  '--accent-secondary': string;
  '--error-color': string;
  '--shadow-sm': string;
  '--shadow-md': string;
  '--shadow-lg': string;
}

interface RoleHours {
  sa: string;
  consultant: string;
  pm: string;
  el: string;
  specialty: string;
}

// Add new interface for Hypercare after RoleHours interface
interface Hypercare {
  hours: string;
  weeks: string;
}

// Add after other interfaces
interface CellMapping {
  sourceId: string;
  targetCell: {
    row: number;
    col: number;
  };
}

interface ButtonProps {
  // ... existing ButtonProps interface ...
}

const lightTheme: Theme = {
  '--bg-primary': '#f8fafc',
  '--bg-secondary': '#ffffff',
  '--bg-tertiary': '#f1f5f9',
  '--text-primary': '#0f172a',
  '--text-secondary': '#334155',
  '--text-tertiary': '#64748b',
  '--accent-primary': '#2563eb',
  '--accent-secondary': '#3b82f6',
  '--error-color': '#ef4444',
  '--shadow-sm': '0 1px 3px rgba(15, 23, 42, 0.1)',
  '--shadow-md': '0 4px 6px -1px rgba(15, 23, 42, 0.1)',
  '--shadow-lg': '0 10px 15px -3px rgba(15, 23, 42, 0.1)',
};

const darkTheme: Theme = {
  '--bg-primary': '#0f172a',
  '--bg-secondary': '#1e293b',
  '--bg-tertiary': '#334155',
  '--text-primary': '#f8fafc',
  '--text-secondary': '#e2e8f0',
  '--text-tertiary': '#cbd5e1',
  '--accent-primary': '#3b82f6',
  '--accent-secondary': '#60a5fa',
  '--error-color': '#f87171',
  '--shadow-sm': '0 1px 3px rgba(0, 0, 0, 0.3)',
  '--shadow-md': '0 4px 6px -1px rgba(0, 0, 0, 0.4)',
  '--shadow-lg': '0 10px 15px -3px rgba(0, 0, 0, 0.4)',
};

const GlobalStyle = styled.div<{ isDarkMode: boolean }>`
  ${({ isDarkMode }) => ({ ...isDarkMode ? darkTheme : lightTheme })}
  background-color: var(--bg-primary);
  color: var(--text-primary);
  min-height: 100vh;
  padding: 2rem;
  transition: all 0.2s ease;
`;

const Container = styled.div`
  max-width: 1400px;
  margin: 0 auto;
  padding: 2rem;
`;

const HeaderContainer = styled.div`
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 2rem;
  position: relative;
`;

const Title = styled.h1`
  color: var(--text-primary);
  margin: 0;
  position: absolute;
  left: 50%;
  transform: translateX(-50%);
  text-align: center;
  font-size: 2rem;
  font-weight: bold;
`;

const HeaderActions = styled.div`
  display: flex;
  gap: 0.5rem;
  margin-left: auto;
`;

const NavButton = styled.button`
  background: none;
  border: none;
  cursor: pointer;
  padding: 0.5rem;
  border-radius: 0.5rem;
  color: var(--text-primary);
  display: flex;
  align-items: center;
  justify-content: center;
  transition: all 0.2s ease;

  &:hover {
    background-color: var(--bg-secondary);
    transform: scale(1.1);
  }

  svg {
    width: 24px;
    height: 24px;
  }
`;

const ThemeToggleButton = styled.button`
  grid-column: 3;
  justify-self: end;
  background: none;
  border: none;
  cursor: pointer;
  padding: 0.5rem;
  border-radius: 0.5rem;
  color: var(--text-primary);
  display: flex;
  align-items: center;
  justify-content: center;
  transition: all 0.2s ease;

  &:hover {
    background-color: var(--bg-secondary);
    transform: scale(1.1);
  }

  svg {
    width: 24px;
    height: 24px;
  }
`;

const SunIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" stroke="none">
    <path d="M12 3a1 1 0 0 1 1 1v1a1 1 0 1 1-2 0V4a1 1 0 0 1 1-1zm0 15a1 1 0 0 1 1 1v1a1 1 0 1 1-2 0v-1a1 1 0 0 1 1-1zm9-7a1 1 0 0 1-1 1h-1a1 1 0 1 1 0-2h1a1 1 0 0 1 1 1zM4 12a1 1 0 0 1-1 1H2a1 1 0 1 1 0-2h1a1 1 0 0 1 1 1zm15.7-7.7a1 1 0 0 1 0 1.4l-1 1a1 1 0 0 1-1.4-1.4l1-1a1 1 0 0 1 1.4 0zM6.7 18.7a1 1 0 0 1 0 1.4l-1 1a1 1 0 0 1-1.4-1.4l1-1a1 1 0 0 1 1.4 0zM18.7 17.3a1 1 0 0 1-1.4 1.4l-1-1a1 1 0 0 1 1.4-1.4l1 1zM7.3 7.3a1 1 0 0 1-1.4 0l-1-1a1 1 0 0 1 1.4-1.4l1 1a1 1 0 0 1 0 1.4zM12 7a5 5 0 1 1 0 10 5 5 0 0 1 0-10z"/>
  </svg>
);

const MoonIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" stroke="none">
    <path d="M12.1 22c-5.5 0-10-4.5-10-10s4.5-10 10-10c.2 0 .5 0 .7.1C10.5 3.7 9 6.7 9 10c0 5 4 9 9 9 .6 0 1.2-.1 1.8-.2-.9 1.9-2.8 3.2-5 3.2h-2.7z"/>
  </svg>
);

const SettingsIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <path d="M12 15a3 3 0 100-6 3 3 0 000 6z" />
    <path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-2 2 2 2 0 01-2-2v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83 0 2 2 0 010-2.83l.06-.06a1.65 1.65 0 00.33-1.82 1.65 1.65 0 00-1.51-1H3a2 2 0 01-2-2 2 2 0 012-2h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 010-2.83 2 2 0 012.83 0l.06.06a1.65 1.65 0 001.82.33H9a1.65 1.65 0 001-1.51V3a2 2 0 012-2 2 2 0 012 2v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 0 2 2 0 010 2.83l-.06.06a1.65 1.65 0 00-.33 1.82V9a1.65 1.65 0 001.51 1H21a2 2 0 012 2 2 2 0 01-2 2h-.09a1.65 1.65 0 00-1.51 1z" />
  </svg>
);

const DocumentIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
    <path d="M14 2v6h6" />
    <line x1="16" y1="13" x2="8" y2="13" />
    <line x1="16" y1="17" x2="8" y2="17" />
    <line x1="10" y1="9" x2="8" y2="9" />
  </svg>
);

const RefreshIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <path d="M21 2v6h-6"></path>
    <path d="M3 12a9 9 0 0 1 15-6.7L21 8"></path>
    <path d="M3 22v-6h6"></path>
    <path d="M21 12a9 9 0 0 1-15 6.7L3 16"></path>
  </svg>
);

const ColumnsContainer = styled.div`
  background: var(--bg-secondary);
  padding: 1rem;
  border-radius: 0.75rem;
  box-shadow: var(--shadow-md);
  margin-bottom: 1.5rem;
`;

const HeaderRow = styled.div`
  display: grid;
  grid-template-columns: 80px 2fr 1.5fr 1.5fr 100px 1.5fr;
  gap: 1rem;
  margin-bottom: 1.25rem;
  padding: 0 0.5rem;
  min-height: 32px;
`;

const ColumnHeader = styled.h3`
  color: var(--text-primary);
  margin: 0;
  text-align: center;
  padding: 0.25rem;
  background-color: var(--bg-tertiary);
  border-radius: 0.5rem;
  font-size: 0.875rem;
  text-transform: uppercase;
  font-weight: 700;
  letter-spacing: 0.05em;
  line-height: 1;
  height: 100%;
  display: flex;
  align-items: center;
  justify-content: center;
`;

const RowsContainer = styled.div`
  display: flex;
  flex-direction: column;
  gap: 0.5rem;
  padding: 0 0.5rem;
`;

const RowWrapper = styled.div`
  display: grid;
  grid-template-columns: 80px 2fr 1.5fr 1.5fr 100px 1.5fr;
  gap: 1rem;
  background: var(--bg-secondary);
  position: relative;
  transition: transform 0.2s ease;
  border-radius: 0.5rem;

  &.selected {
    background: var(--bg-tertiary);
  }

  &.dragging {
    opacity: 0.9;
    box-shadow: var(--shadow-lg);
    z-index: 20;
  }
`;

const CheckboxContainer = styled.div`
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 0.25rem 0;
  height: 100%;
`;

const HeaderCheckboxContainer = styled.div`
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 0.25rem 0;
`;

const Checkbox = styled.input`
  width: 1.25rem;
  height: 1.25rem;
  cursor: pointer;
  border: 2px solid var(--text-tertiary);
  border-radius: 0.25rem;
  
  &:checked {
    background-color: var(--accent-primary);
    border-color: var(--accent-primary);
  }
`;

const DragHandle = styled.div`
  cursor: grab;
  color: var(--text-tertiary);
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 24px;
  width: 32px;
  height: 32px;
  border-radius: 6px;
  transition: all 0.2s ease;
  margin-right: 0.5rem;
  position: relative;
  overflow: hidden;

  &::before {
    content: '';
    position: absolute;
    inset: 0;
    background: var(--accent-primary);
    opacity: 0;
    transition: opacity 0.2s ease;
  }

  &:hover {
    color: var(--accent-primary);
    background: var(--bg-tertiary);
    transform: translateY(-1px);
    
    &::before {
      opacity: 0.1;
    }
  }

  &:active {
    cursor: grabbing;
    transform: translateY(0);
    background: var(--bg-tertiary);
    
    &::before {
      opacity: 0.2;
    }
  }

  svg {
    width: 20px;
    height: 20px;
    position: relative;
    z-index: 1;
  }
`;

const NumberInputContainer = styled.div`
  width: 100%;
  display: flex;
  align-items: center;
  justify-content: center;
`;

const NumberInput = styled.input`
  width: 100%;
  padding: 0.5rem;
  border: 1px solid var(--text-tertiary);
  border-radius: 0.25rem;
  background: var(--bg-primary);
  color: var(--text-primary);
  text-align: center;
  
  &:focus {
    outline: none;
    border-color: var(--accent-primary);
    box-shadow: 0 0 0 2px var(--accent-secondary);
  }
`;

const Button = styled.button<ButtonProps>`
  padding: 0.5rem 1rem;
  border: none;
  border-radius: 0.5rem;
  background-color: var(--accent-primary);
  color: white;
  font-weight: 600;
  cursor: pointer;
  transition: all 0.2s ease;
  display: flex;
  align-items: center;
  gap: 0.5rem;
  
  &:hover {
    background-color: var(--accent-secondary);
  }
  
  &:disabled {
    opacity: 0.5;
    cursor: not-allowed;
  }

  svg {
    width: 16px;
    height: 16px;
  }
`;

const MenuBar = styled.div`
  position: absolute;
  top: -56px;
  left: 0;
  right: 0;
  display: flex;
  gap: 0.5rem;
  padding: 0.5rem;
  min-width: fit-content;
  width: 100%;
  background-color: var(--bg-primary);
  border: 1px solid var(--text-tertiary);
  border-radius: 0.5rem;
  box-shadow: var(--shadow-md);
  z-index: 30;
  opacity: 0;
  transform: translateY(10px);
  transition: all 0.2s ease;
  pointer-events: none;

  &.visible {
    opacity: 1;
    transform: translateY(0);
    pointer-events: all;
  }
`;

const MenuButton = styled.button<ButtonProps>`
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 32px;
  height: 32px;
  padding: 0;
  border: none;
  border-radius: 0.25rem;
  background: transparent;
  color: var(--text-primary);
  font-size: 1rem;
  cursor: pointer;
  transition: all 0.2s ease;

  &:hover {
    background-color: var(--bg-secondary);
  }

  &.is-active {
    background-color: var(--accent-primary);
    color: white;
  }

  svg {
    width: 16px;
    height: 16px;
  }
`;

const TotalContainer = styled.div`
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-top: 1.5rem;
  padding: 1.5rem 2rem;
  background-color: var(--bg-secondary);
  border-radius: 1rem;
  box-shadow: var(--shadow-md);
  border: 2px solid var(--accent-primary);
  transition: all 0.2s ease;

  &:hover {
    transform: translateY(-2px);
    box-shadow: var(--shadow-lg);
  }
`;

const ActionButtons = styled.div`
  display: flex;
  gap: 1rem;
`;

const DeleteButton = styled(Button)`
  background-color: #dc2626;
  padding: 0.75rem 1.5rem;
  font-size: 1.1rem;
  
  svg {
    width: 20px;
    height: 20px;
  }
  
  &:hover {
    background-color: #b91c1c;
  }

  &:disabled {
    opacity: 0.5;
    background-color: #dc2626;
  }
`;

const DuplicateButton = styled(Button)`
  background-color: #2563eb;
  padding: 0.75rem 1.5rem;
  font-size: 1.1rem;
  
  svg {
    width: 20px;
    height: 20px;
  }
  
  &:hover {
    background-color: #1d4ed8;
  }

  &:disabled {
    opacity: 0.5;
    background-color: #2563eb;
  }
`;

const TotalSection = styled.div`
  display: flex;
  align-items: center;
  gap: 1.5rem;
`;

const TotalLabel = styled.span`
  font-size: 1.25rem;
  font-weight: 700;
  color: var(--text-primary);
  text-transform: uppercase;
  letter-spacing: 0.05em;
`;

const TotalValue = styled.span`
  font-size: 2rem;
  font-weight: 800;
  color: var(--accent-primary);
  padding: 0.5rem 1rem;
  background-color: var(--bg-primary);
  border-radius: 0.75rem;
  min-width: 120px;
  text-align: center;
  box-shadow: var(--shadow-sm);
  
  &::after {
    content: ' build hrs';
    font-size: 1rem;
    color: var(--text-secondary);
    font-weight: 600;
  }
`;

const ExportButton = styled(Button)`
  background-color: #0d9488;
  padding: 0.75rem 1.5rem;
  font-size: 1.1rem;
  
  svg {
    width: 20px;
    height: 20px;
  }
  
  &:hover {
    background-color: #0f766e;
  }
`;

const AddRowButton = styled(Button)`
  background-color: #2563eb;
  padding: 0.75rem 1.5rem;
  font-size: 1.1rem;
  
  svg {
    width: 20px;
    height: 20px;
  }
  
  &:hover {
    background-color: #1d4ed8;
  }
`;

const FloatingActionButton = styled.button`
  position: fixed;
  bottom: 2rem;
  right: 2rem;
  width: 56px;
  height: 56px;
  border-radius: 50%;
  background-color: var(--accent-primary);
  color: white;
  border: none;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 24px;
  box-shadow: var(--shadow-lg);
  transition: all 0.2s ease;
  z-index: 100;

  &:hover {
    background-color: var(--accent-secondary);
    transform: translateY(-2px);
    box-shadow: 0 12px 20px -8px rgba(0, 0, 0, 0.2);
  }

  &:active {
    transform: translateY(0);
  }

  svg {
    width: 24px;
    height: 24px;
  }
`;

const AddIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <line x1="12" y1="5" x2="12" y2="19"></line>
    <line x1="5" y1="12" x2="19" y2="12"></line>
  </svg>
);

const EditorContainer = styled.div`
  position: relative;
  width: 100%;

  .ProseMirror {
    padding: 0.35rem 0.5rem;
    border: 1px solid var(--text-tertiary);
    border-radius: 0.25rem;
    min-height: 80px;
    max-height: 160px;
    overflow-y: auto;
    background: var(--bg-primary);
    color: var(--text-primary);
    overflow-wrap: break-word;
    word-wrap: break-word;
    word-break: break-word;
    transition: all 0.2s ease;
    line-height: 1.3;

    &:focus {
      outline: none;
      border-color: var(--accent-primary);
      box-shadow: 0 0 0 2px var(--accent-secondary);
    }

    p {
      margin: 0;
      white-space: pre-wrap;
    }

    ul, ol {
      margin: 0;
      padding-left: 1.2rem;
    }
  }
`;

// Add new styled components for dropdowns
const Dropdown = styled.div`
  position: relative;
  display: inline-block;
`;

const DropdownButton = styled(MenuButton)`
  display: flex;
  align-items: center;
  gap: 0.25rem;
  min-width: 80px;
  justify-content: center;
`;

const DropdownContent = styled.div`
  position: absolute;
  top: 100%;
  left: 0;
  background-color: var(--bg-primary);
  border: 1px solid var(--text-tertiary);
  border-radius: 0.25rem;
  box-shadow: var(--shadow-md);
  z-index: 1000;
  min-width: 120px;
`;

const DropdownItem = styled.button`
  width: 100%;
  padding: 0.5rem;
  border: none;
  background: none;
  color: var(--text-primary);
  text-align: left;
  cursor: pointer;
  display: flex;
  align-items: center;
  gap: 0.5rem;

  &:hover {
    background-color: var(--bg-secondary);
  }

  &.active {
    background-color: var(--bg-tertiary);
  }
`;

// Icons components
const BoldIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <path d="M6 4h8a4 4 0 0 1 4 4 4 4 0 0 1-4 4H6z"></path>
    <path d="M6 12h9a4 4 0 0 1 4 4 4 4 0 0 1-4 4H6z"></path>
  </svg>
);

const ItalicIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <line x1="19" y1="4" x2="10" y2="4"></line>
    <line x1="14" y1="20" x2="5" y2="20"></line>
    <line x1="15" y1="4" x2="9" y2="20"></line>
  </svg>
);

const RichTextEditor = ({ content, onChange }: RichTextEditorProps) => {
  const [showListMenu, setShowListMenu] = useState(false);
  const [showAlignMenu, setShowAlignMenu] = useState(false);
  const [showColorMenu, setShowColorMenu] = useState(false);
  const [isFocused, setIsFocused] = useState(false);
  const editorContainerRef = useRef<HTMLDivElement>(null);
  
  const editor = useEditor({
    extensions: [
      StarterKit,
      TextStyle,
      Color,
      TextAlign.configure({
        types: ['heading', 'paragraph'],
      }),
    ],
    content,
    onUpdate: ({ editor }) => {
      onChange(editor.getHTML());
    },
    onFocus: () => {
      setIsFocused(true);
    },
    onBlur: () => {
      // Don't hide toolbar if clicking within the editor container
      if (editorContainerRef.current?.contains(document.activeElement)) {
        return;
      }
      setTimeout(() => {
        if (!editorContainerRef.current?.contains(document.activeElement)) {
          setIsFocused(false);
          closeAllDropdowns();
        }
      }, 100);
    },
  });

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (editorContainerRef.current && !editorContainerRef.current.contains(event.target as Node)) {
        setIsFocused(false);
        closeAllDropdowns();
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, []);

  if (!editor) {
    return null;
  }

  const closeAllDropdowns = () => {
    setShowListMenu(false);
    setShowAlignMenu(false);
    setShowColorMenu(false);
  };

  return (
    <EditorContainer ref={editorContainerRef}>
      <MenuBar className={isFocused ? 'visible' : ''}>
        <MenuButton
          onClick={() => editor.chain().focus().toggleBold().run()}
          className={editor.isActive('bold') ? 'is-active' : ''}
        >
          <BoldIcon />
        </MenuButton>
        <MenuButton
          onClick={() => editor.chain().focus().toggleItalic().run()}
          className={editor.isActive('italic') ? 'is-active' : ''}
        >
          <ItalicIcon />
        </MenuButton>

        <Dropdown>
          <DropdownButton
            onClick={() => {
              closeAllDropdowns();
              setShowListMenu(!showListMenu);
            }}
          >
            List ▾
          </DropdownButton>
          {showListMenu && (
            <DropdownContent>
              <DropdownItem
                onClick={() => {
                  editor.chain().focus().toggleBulletList().run();
                  setShowListMenu(false);
                }}
                className={editor.isActive('bulletList') ? 'active' : ''}
              >
                • Bullet List
              </DropdownItem>
              <DropdownItem
                onClick={() => {
                  editor.chain().focus().toggleOrderedList().run();
                  setShowListMenu(false);
                }}
                className={editor.isActive('orderedList') ? 'active' : ''}
              >
                1. Numbered List
              </DropdownItem>
            </DropdownContent>
          )}
        </Dropdown>

        <Dropdown>
          <DropdownButton
            onClick={() => {
              closeAllDropdowns();
              setShowAlignMenu(!showAlignMenu);
            }}
          >
            Align ▾
          </DropdownButton>
          {showAlignMenu && (
            <DropdownContent>
              <DropdownItem
                onClick={() => {
                  editor.chain().focus().setTextAlign('left').run();
                  setShowAlignMenu(false);
                }}
                className={editor.isActive({ textAlign: 'left' }) ? 'active' : ''}
              >
                ⇤ Left
              </DropdownItem>
              <DropdownItem
                onClick={() => {
                  editor.chain().focus().setTextAlign('center').run();
                  setShowAlignMenu(false);
                }}
                className={editor.isActive({ textAlign: 'center' }) ? 'active' : ''}
              >
                ⇔ Center
              </DropdownItem>
              <DropdownItem
                onClick={() => {
                  editor.chain().focus().setTextAlign('right').run();
                  setShowAlignMenu(false);
                }}
                className={editor.isActive({ textAlign: 'right' }) ? 'active' : ''}
              >
                ⇥ Right
              </DropdownItem>
            </DropdownContent>
          )}
        </Dropdown>

        <Dropdown>
          <DropdownButton
            onClick={() => {
              closeAllDropdowns();
              setShowColorMenu(!showColorMenu);
            }}
          >
            Color ▾
          </DropdownButton>
          {showColorMenu && (
            <DropdownContent>
              <DropdownItem
                onClick={() => {
                  editor.chain().focus().unsetColor().run();
                  setShowColorMenu(false);
                }}
              >
                <span style={{ color: 'var(--text-primary)' }}>⬤</span> Default
              </DropdownItem>
              <DropdownItem
                onClick={() => {
                  editor.chain().focus().setColor('#EF4444').run();
                  setShowColorMenu(false);
                }}
              >
                <span style={{ color: '#EF4444' }}>⬤</span> Red
              </DropdownItem>
              <DropdownItem
                onClick={() => {
                  editor.chain().focus().setColor('#3B82F6').run();
                  setShowColorMenu(false);
                }}
              >
                <span style={{ color: '#3B82F6' }}>⬤</span> Blue
              </DropdownItem>
            </DropdownContent>
          )}
        </Dropdown>
      </MenuBar>
      <EditorContent editor={editor} />
    </EditorContainer>
  );
};

const DragHandleIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor">
    <path d="M8 4a1.5 1.5 0 100-3 1.5 1.5 0 000 3zm8 0a1.5 1.5 0 100-3 1.5 1.5 0 000 3zM8 11a1.5 1.5 0 100-3 1.5 1.5 0 000 3zm8 0a1.5 1.5 0 100-3 1.5 1.5 0 000 3zm-8 7a1.5 1.5 0 100-3 1.5 1.5 0 000 3zm8 0a1.5 1.5 0 100-3 1.5 1.5 0 000 3z" />
  </svg>
);

const SortableRow = ({ id, index, row, isSelected, onToggleSelect, onUpdateRow }: SortableRowProps) => {
  const {
    attributes,
    listeners,
    setNodeRef,
    transform,
    transition,
    isDragging,
  } = useSortable({ id });

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
    zIndex: isDragging ? 1 : 0,
  };

  return (
    <RowWrapper
      ref={setNodeRef}
      style={style}
      className={`${isDragging ? 'dragging' : ''} ${isSelected ? 'selected' : ''}`}
    >
      <CheckboxContainer>
        <div {...attributes} {...listeners}>
          <DragHandle>
            <DragHandleIcon />
          </DragHandle>
        </div>
        <Checkbox
          type="checkbox"
          checked={isSelected}
          onChange={onToggleSelect}
        />
      </CheckboxContainer>

      <RichTextEditor
        content={row.processAndImpact}
        onChange={(value) => onUpdateRow('processAndImpact', value)}
      />

      <RichTextEditor
        content={row.components}
        onChange={(value) => onUpdateRow('components', value)}
      />

      <RichTextEditor
        content={row.assumptions}
        onChange={(value) => onUpdateRow('assumptions', value)}
      />

      <NumberInputContainer>
        <NumberInput
          type="number"
          min="0"
          step="0.5"
          value={row.hours}
          onChange={(e) => onUpdateRow('hours', e.target.value)}
        />
      </NumberInputContainer>

      <RichTextEditor
        content={row.notes}
        onChange={(value) => onUpdateRow('notes', value)}
      />
    </RowWrapper>
  );
};

// Add ConfirmationModal component
const ModalOverlay = styled.div`
  position: fixed;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background-color: rgba(0, 0, 0, 0.5);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 1000;
`;

const ModalContent = styled.div`
  background: var(--bg-primary);
  padding: 2rem;
  border-radius: 0.75rem;
  box-shadow: var(--shadow-lg);
  max-width: 500px;
  width: 90%;
`;

const ModalTitle = styled.h2`
  color: var(--text-primary);
  margin: 0 0 1rem 0;
  font-size: 1.5rem;
`;

const ModalMessage = styled.p`
  color: var(--text-secondary);
  margin-bottom: 2rem;
`;

const ModalButtons = styled.div`
  display: flex;
  justify-content: flex-end;
  gap: 1rem;
`;

const ConfirmationModal = ({ isOpen, onClose, onConfirm, title, message, confirmText = 'Confirm', cancelText = 'Cancel' }: ConfirmationModalProps) => {
  if (!isOpen) return null;

  return (
    <ModalOverlay onClick={onClose}>
      <ModalContent onClick={e => e.stopPropagation()}>
        <ModalTitle>{title}</ModalTitle>
        <ModalMessage>{message}</ModalMessage>
        <ModalButtons>
          <Button onClick={onClose}>{cancelText}</Button>
          <Button 
            onClick={onConfirm}
            style={{ backgroundColor: 'var(--accent-primary)' }}
          >
            {confirmText}
          </Button>
        </ModalButtons>
      </ModalContent>
    </ModalOverlay>
  );
};

// Add new icon components
const DeleteIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <path d="M3 6h18"></path>
    <path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"></path>
    <path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"></path>
    <line x1="10" y1="11" x2="10" y2="17"></line>
    <line x1="14" y1="11" x2="14" y2="17"></line>
  </svg>
);

const DuplicateIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
    <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
  </svg>
);

const ThemeIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <circle cx="12" cy="12" r="5"></circle>
    <path d="M12 1v2"></path>
    <path d="M12 21v2"></path>
    <path d="M4.22 4.22l1.42 1.42"></path>
    <path d="M18.36 18.36l1.42 1.42"></path>
    <path d="M1 12h2"></path>
    <path d="M21 12h2"></path>
    <path d="M4.22 19.78l1.42-1.42"></path>
    <path d="M18.36 5.64l1.42-1.42"></path>
  </svg>
);

const UploadContainer = styled.div`
  max-width: 600px;
  margin: 2rem auto;
  padding: 1rem;
`;

const DropZone = styled.div<{ isDragging: boolean; isError?: boolean }>`
  border: 2px dashed ${({ theme, isDragging, isError }) => 
    isError ? 'var(--error-color)' : 
    isDragging ? 'var(--accent-primary)' : 'var(--text-tertiary)'};
  border-radius: 8px;
  padding: 2rem;
  text-align: center;
  background: ${({ theme, isDragging }) => 
    isDragging ? 'var(--bg-tertiary)' : 'var(--bg-secondary)'};
  transition: all 0.2s ease;
  cursor: pointer;
  
  &:hover {
    border-color: var(--accent-primary);
    background: var(--bg-tertiary);
  }
`;

const FileInput = styled.input`
  display: none;
`;

const UploadMessage = styled.p`
  margin: 0;
  color: var(--text-primary);
  font-size: 1rem;
`;

const FilePreview = styled.div`
  margin-top: 1rem;
  padding: 1rem;
  background: var(--bg-secondary);
  border: 1px solid var(--text-tertiary);
  border-radius: 4px;
  display: flex;
  align-items: center;
  justify-content: space-between;
`;

const ErrorMessage = styled.p`
  color: var(--error-color, #ef4444);
  margin: 0.5rem 0;
  font-size: 0.9rem;
`;

interface TemplateFile {
  name: string;
  size: number;
  type: string;
}

const emptyRow: RowData = {
  id: crypto.randomUUID(),
  processAndImpact: '',
  components: '',
  assumptions: '',
  hours: '',
  notes: ''
};

// Add new interfaces
interface ColumnMapping {
  processAndImpact: string;
  components: string;
  assumptions: string;
  hours: string;
  notes: string;
}

interface TemplateData {
  columns: string[];
  mapping: ColumnMapping | null;
  rows: any[]; // Add this to store the actual Excel data
}

// Add new styled components after existing ones
const MappingContainer = styled.div`
  margin-top: 2rem;
  padding: 1.5rem;
  background: var(--bg-secondary);
  border-radius: 0.75rem;
  box-shadow: var(--shadow-md);
`;

const MappingTitle = styled.h3`
  color: var(--text-primary);
  margin: 0 0 1.5rem 0;
  font-size: 1.25rem;
`;

const MappingRow = styled.div`
  display: grid;
  grid-template-columns: 1fr 2fr;
  gap: 1rem;
  margin-bottom: 1rem;
  align-items: center;
`;

const MappingLabel = styled.label`
  color: var(--text-primary);
  font-weight: 600;
`;

const MappingSelect = styled.select`
  width: 100%;
  padding: 0.5rem;
  border: 1px solid var(--text-tertiary);
  border-radius: 0.25rem;
  background: var(--bg-primary);
  color: var(--text-primary);
  font-size: 1rem;
  
  &:focus {
    outline: none;
    border-color: var(--accent-primary);
    box-shadow: 0 0 0 2px var(--accent-secondary);
  }
`;

const SaveButton = styled(Button)`
  margin-top: 1.5rem;
  background-color: #15803d;
  
  &:hover {
    background-color: #166534;
  }

  &:disabled {
    opacity: 0.5;
    cursor: not-allowed;
  }
`;

const ApplyButton = styled(Button)`
  margin-top: 1rem;
  background-color: #2563eb;
  
  &:hover {
    background-color: #1d4ed8;
  }

  &:disabled {
    opacity: 0.5;
    cursor: not-allowed;
  }
`;

const SaveMessage = styled.div<{ success?: boolean }>`
  margin-top: 1rem;
  padding: 0.75rem;
  border-radius: 0.375rem;
  background-color: ${({ success }) => 
    success ? 'var(--bg-tertiary)' : 'var(--error-color)'};
  color: ${({ success }) => 
    success ? 'var(--text-primary)' : 'white'};
  font-size: 0.875rem;
  display: flex;
  align-items: center;
  gap: 0.5rem;
`;

// Add new styled components after existing ones
const PreviewContainer = styled.div`
  margin-top: 2rem;
  overflow-x: auto;
  background: var(--bg-secondary);
  border-radius: 0.75rem;
  box-shadow: var(--shadow-md);
`;

const PreviewTable = styled.table`
  width: 100%;
  border-collapse: collapse;
  font-size: 0.875rem;
`;

const PreviewHeader = styled.th<{ isSelected?: boolean }>`
  padding: 0.75rem;
  background: var(--bg-tertiary);
  border: 1px solid var(--text-tertiary);
  color: var(--text-primary);
  font-weight: 600;
  text-align: left;
  cursor: pointer;
  position: relative;
  transition: all 0.2s ease;

  ${({ isSelected }) => isSelected && `
    background: var(--accent-primary);
    color: white;
  `}

  &:hover {
    background: ${({ isSelected }) => 
      isSelected ? 'var(--accent-secondary)' : 'var(--bg-primary)'};
  }
`;

const PreviewCell = styled.td`
  padding: 0.5rem 0.75rem;
  border: 1px solid var(--text-tertiary);
  color: var(--text-secondary);
  max-width: 200px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
`;

const MappingOverlay = styled.div`
  position: fixed;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background: rgba(0, 0, 0, 0.5);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 1000;
`;

const MappingDialog = styled.div`
  background: var(--bg-primary);
  padding: 1.5rem;
  border-radius: 0.75rem;
  box-shadow: var(--shadow-lg);
  max-width: 400px;
  width: 90%;
`;

const MappingOptions = styled.div`
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 1rem;
  margin: 1rem 0;
`;

const MappingOption = styled.button<{ isSelected?: boolean }>`
  padding: 1rem;
  border: 2px solid ${({ isSelected }) => 
    isSelected ? 'var(--accent-primary)' : 'var(--text-tertiary)'};
  border-radius: 0.5rem;
  background: var(--bg-secondary);
  color: var(--text-primary);
  cursor: pointer;
  transition: all 0.2s ease;
  font-weight: ${({ isSelected }) => isSelected ? '600' : '400'};

  &:hover {
    border-color: var(--accent-primary);
    background: var(--bg-tertiary);
  }
`;

// Add helper functions
const stripHtml = (html: string) => {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  return doc.body.textContent || '';
};

const URLInput = styled.input`
  width: 100%;
  padding: 0.75rem;
  border: 2px solid var(--text-tertiary);
  border-radius: 0.5rem;
  background: var(--bg-secondary);
  color: var(--text-primary);
  font-size: 1rem;
  transition: all 0.2s ease;

  &:focus {
    outline: none;
    border-color: var(--accent-primary);
  }

  &::placeholder {
    color: var(--text-tertiary);
  }
`;

const Instructions = styled.p`
  color: var(--text-secondary);
  font-size: 0.875rem;
  margin: 0.5rem 0 1.5rem;
  line-height: 1.5;
`;

interface GoogleSheetsConfig {
  apiKey: string;
  spreadsheetId: string;
  range: string;
}

// Add this interface before the SheetData interface
interface TextFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontSize?: number;
  foregroundColor?: {
    red: number;
    green: number;
    blue: number;
    alpha?: number;
  };
}

// Add these interfaces before the SheetData interface
interface SheetProperties {
  sheetId: number;
  title: string;
  gridProperties: {
    rowCount: number;
    columnCount: number;
  };
}

interface Sheet {
  properties: SheetProperties;
  data: {
    rowMetadata?: { pixelSize: number }[];
    columnMetadata?: { pixelSize: number }[];
    rowData?: {
      values?: {
        userEnteredFormat?: {
          textFormat?: TextFormat;
          backgroundColor?: {
            red: number;
            green: number;
            blue: number;
            alpha?: number;
          };
          horizontalAlignment?: 'LEFT' | 'CENTER' | 'RIGHT' | 'JUSTIFY';
          verticalAlignment?: 'TOP' | 'MIDDLE' | 'BOTTOM';
        };
      }[];
    }[];
  }[];
}

interface SheetData {
  headers: string[];
  values: string[][];
  outOfScopeItems?: string[][];
  formatting?: {
    values?: {
      userEnteredFormat?: {
        textFormat?: TextFormat;
        backgroundColor?: {
          red: number;
          green: number;
          blue: number;
          alpha?: number;
        };
        horizontalAlignment?: 'LEFT' | 'CENTER' | 'RIGHT' | 'JUSTIFY';
        verticalAlignment?: 'TOP' | 'MIDDLE' | 'BOTTOM';
      };
    }[];
  }[];
  columnMetadata?: {
    pixelSize: number;
  }[];
  rowMetadata?: {
    pixelSize: number;
  }[];
}

const ConfigSection = styled.div`
  margin-bottom: 2rem;
`;

const Input = styled.input`
  width: 100%;
  padding: 0.5rem;
  border: 1px solid var(--bg-tertiary);
  border-radius: 0.5rem;
  background: var(--bg-secondary);
  color: var(--text-primary);
  margin-bottom: 1rem;
  
  &:focus {
    outline: none;
    border-color: var(--accent-primary);
  }
`;

const Label = styled.label`
  display: block;
  margin-bottom: 0.5rem;
  color: var(--text-secondary);
  font-size: 0.875rem;
`;

const SheetPreview = styled.div`
  background: var(--bg-secondary);
  border-radius: 0.75rem;
  padding: 0;  // Remove padding
  box-shadow: var(--shadow-md);
  margin-top: 1rem;
  max-width: 100%;
  max-height: 90vh;
  display: flex;
  flex-direction: column;
`;

const SheetGridContainer = styled.div`
  flex: 1;
  overflow: auto;
  border: 1px solid #e5e7eb;
  border-radius: 0.75rem;  // Match parent's border radius
  background: white;
  min-height: 0; /* Important for flex child scrolling */
`;

const SheetGrid = styled.div`
  display: table;
  width: 100%;
  background: white;
  border-collapse: collapse;
  font-family: 'Roboto', sans-serif;
  table-layout: fixed;
  min-width: max-content;
  
  &::-webkit-scrollbar {
    width: 10px;
    height: 10px;
  }
  
  &::-webkit-scrollbar-track {
    background: #f1f1f1;
    border-radius: 5px;
  }
  
  &::-webkit-scrollbar-thumb {
    background: #c1c1c1;
    border-radius: 5px;
    border: 2px solid #f1f1f1;
    
    &:hover {
      background: #a1a1a1;
    }
  }
`;

const SheetRow = styled.div<{ height?: string }>`
  display: table-row;
  height: ${props => props.height || '21px'}; // Default Google Sheets row height
`;

const CornerCell = styled.div`
  display: table-cell;
  width: 40px;
  background: #f1f3f4;
  border: 1px solid #e5e7eb;
  position: sticky;
  left: 0;
  top: 0;
  z-index: 20;
`;

const RowHeader = styled.div`
  display: table-cell;
  width: 40px;
  background: #f1f3f4;
  border: 1px solid #e5e7eb;
  font-size: 11px;
  line-height: 15px;
  color: #5f6368;
  text-align: center;
  position: sticky;
  left: 0;
  z-index: 10;
`;

const SheetCell = styled.div<{ 
  isHeader?: boolean; 
  isSelected?: boolean;
  isEditing?: boolean;
  width?: string;
  backgroundColor?: string;
  textColor?: string;
  textAlign?: 'left' | 'center' | 'right';
  fontSize?: string;
  isHighlighted?: boolean;
  isMapped?: boolean;
}>`
  display: table-cell;
  padding: 3px 6px;
  border: 1px solid #e5e7eb;
  background: ${props => 
    props.isMapped ? 'rgba(37, 99, 235, 0.1)' :  // Light blue background for mapped cells
    props.backgroundColor ? props.backgroundColor :
    props.isEditing ? '#e8f0fe' :
    props.isSelected ? '#e8f0fe' :
    props.isHeader ? '#f1f3f4' : 'white'};
  font-size: ${props => props.fontSize || '11px'};
  line-height: 1.4;
  color: ${props => props.textColor ? props.textColor : props.isHeader ? '#5f6368' : '#000'};
  position: relative;
  vertical-align: middle;
  width: ${props => props.width || '100px'};
  min-width: ${props => props.width || '100px'};
  max-width: ${props => props.width || '100px'};
  white-space: pre-wrap;
  overflow: hidden;
  text-overflow: ellipsis;
  text-align: ${props => props.textAlign || (props.isHeader ? 'center' : 'left')};
  
  ${props => props.isHeader && `
    font-weight: 500;
    position: sticky;
    top: 0;
    z-index: 10;
  `}
  
  &:hover {
    background: ${props => 
      props.isMapped ? 'rgba(37, 99, 235, 0.15)' :  // Slightly darker on hover
      props.isHeader ? '#e8eaed' : '#f8f9fa'};
  }

  ${props => props.isHighlighted && !props.isMapped && `
    background-color: var(--accent-secondary) !important;
    opacity: 0.7;
  `}

  ${props => props.isMapped && `
    border: 2px solid rgba(37, 99, 235, 0.3);
    padding-top: 12px; /* Make room for the field title */

    &::before {
      content: attr(data-field-title);
      position: absolute;
      top: 2px;
      left: 6px;
      font-size: 9px;
      color: rgba(0, 0, 0, 0.4);
      font-weight: 500;
      text-transform: uppercase;
      letter-spacing: 0.02em;
    }

    .delete-mapping {
      position: absolute;
      top: 2px;
      right: 2px;
      width: 12px;
      height: 12px;
      padding: 0;
      display: flex;
      align-items: center;
      justify-content: center;
      border: none;
      background: none;
      color: var(--text-tertiary);
      cursor: pointer;
      opacity: 0.6;
      transition: all 0.2s ease;

      &:hover {
        opacity: 1;
        color: var(--error-color);
      }

      svg {
        width: 10px;
        height: 10px;
      }
    }
  `}
`;

const CellInput = styled.textarea`
  position: absolute;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  padding: 3px 6px;
  border: 2px solid #1a73e8;
  background: white;
  font-size: 11px;
  line-height: 15px;
  font-family: inherit;
  outline: none;
  z-index: 20;
  resize: none;
  overflow: hidden;
`;

const LoadingOverlay = styled.div`
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background: rgba(255, 255, 255, 0.8);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 30;
  font-size: 14px;
  color: #1a73e8;
`;

interface CellPosition {
  row: number;
  col: number;
}

const RefreshButton = styled(Button)`
  background-color: var(--accent-primary);
  padding: 0.5rem;
  margin-left: 1rem;
  
  svg {
    width: 20px;
    height: 20px;
  }
  
  &:hover {
    background-color: var(--accent-secondary);
    transform: scale(1.05);
  }

  &:disabled {
    opacity: 0.5;
    cursor: not-allowed;
  }
`;

const GoogleSheetsPreview: React.FC<{
  config: GoogleSheetsConfig;
  onDataLoaded: (data: SheetData) => void;
  setLastMapping: (mapping: CellMapping | null) => void;
  setCellMappings: React.Dispatch<React.SetStateAction<CellMapping[]>>;
  cellMappings: CellMapping[];
}> = ({ config, onDataLoaded, setLastMapping, setCellMappings, cellMappings }) => {
  const [sheetData, setSheetData] = useState<SheetData | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [selectedCell, setSelectedCell] = useState<CellPosition | null>(null);
  const [editingCell, setEditingCell] = useState<CellPosition | null>(null);
  const [editValue, setEditValue] = useState('');
  const [columnWidths, setColumnWidths] = useState<string[]>([]);
  const [dragOverCell, setDragOverCell] = useState<CellPosition | null>(null);

  const fetchSheetData = async () => {
    try {
      setLoading(true);
      
      // First, fetch the spreadsheet metadata to get the sheet name
      const spreadsheetResponse = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${config.spreadsheetId}?key=${config.apiKey}`
      );
      
      if (!spreadsheetResponse.ok) {
        const errorData = await spreadsheetResponse.json();
        throw new Error(errorData.error?.message || 'Failed to fetch spreadsheet metadata');
      }

      const spreadsheetData = await spreadsheetResponse.json();
      const firstSheet = spreadsheetData.sheets?.[0];
      const sheetTitle = firstSheet?.properties?.title || 'Sheet1';
      
      // Use the actual sheet name in the range
      const range = `${sheetTitle}!A1:Z1000`;
      
      // Fetch values with the correct sheet name
      const valuesResponse = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${config.spreadsheetId}/values/${encodeURIComponent(range)}?key=${config.apiKey}`
      );
      
      if (!valuesResponse.ok) {
        const errorData = await valuesResponse.json();
        throw new Error(errorData.error?.message || 'Failed to fetch sheet data');
      }

      const valuesData = await valuesResponse.json();

      // Fetch sheet metadata including dimensions
      const metadataResponse = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${config.spreadsheetId}?key=${config.apiKey}&fields=sheets(properties(sheetId,title,gridProperties),data(rowMetadata/pixelSize,columnMetadata/pixelSize,rowData/values/userEnteredFormat))`
      );

      if (!metadataResponse.ok) {
        const errorData = await metadataResponse.json();
        throw new Error(errorData.error?.message || 'Failed to fetch metadata');
      }

      const metadataData = await metadataResponse.json();
      const sheetMetadata = metadataData.sheets?.find((sheet: Sheet) => sheet.properties?.title === sheetTitle) || metadataData.sheets?.[0];

      if (!sheetMetadata) {
        throw new Error('No sheet metadata found');
      }
      
      // Get the maximum number of columns from all rows
      const maxColumns = Math.max(...valuesData.values.map((row: string[]) => row.length));
      
      // Normalize all rows to have the same number of columns
      const normalizedValues = valuesData.values.map((row: string[]) => {
        const paddedRow = [...row];
        while (paddedRow.length < maxColumns) {
          paddedRow.push('');
        }
        return paddedRow;
      });

      // Get column widths from metadata or calculate based on content
      const columnMetadata = sheetMetadata.data?.[0]?.columnMetadata || [];
      const widths = Array(maxColumns).fill('').map((_, colIndex) => {
        // If we have metadata for this column, use it
        if (columnMetadata[colIndex]?.pixelSize) {
          return `${columnMetadata[colIndex].pixelSize}px`;
        }
        
        // Otherwise calculate based on content
        const columnContent = normalizedValues.map((row: string[]) => row[colIndex] || '');
        const maxContentLength = Math.max(...columnContent.map((content: string) => 
          content.toString().length
        ));
        const calculatedWidth = Math.min(Math.max(maxContentLength * 8, 100), 300);
        return `${calculatedWidth}px`;
      });
      setColumnWidths(widths);

      const headers = normalizedValues[0] || Array(maxColumns).fill('');
      const values = normalizedValues.slice(1) || [];
      
      interface SheetRowValue {
        userEnteredFormat?: {
          textFormat?: {
            strikethrough?: boolean;
          };
        };
      }

      interface SheetRow {
        values?: SheetRowValue[];
      }
      
      const sheetData = { 
        headers, 
        values,
        formatting: sheetMetadata.data?.[0]?.rowData || [],
        columnMetadata: sheetMetadata.data?.[0]?.columnMetadata || [],
        rowMetadata: sheetMetadata.data?.[0]?.rowMetadata || [],
        outOfScopeItems: sheetMetadata.data?.[0]?.rowData?.find((row: SheetRow) => 
          row.values?.some((value: SheetRowValue) => 
            value.userEnteredFormat?.textFormat?.strikethrough
          )
        )?.values?.slice(1) || [['']]
      };
      
      setSheetData(sheetData);
      onDataLoaded(sheetData);
      setError(null);
    } catch (err) {
      console.error('Fetch error:', err);
      setError(err instanceof Error ? err.message : 'An error occurred while fetching the sheet data');
      setSheetData(null);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (config.apiKey && config.spreadsheetId && config.range) {
      fetchSheetData();
    }
  }, [config, onDataLoaded]);

  const handleCellClick = (row: number, col: number) => {
    setSelectedCell({ row, col });
    if (row > 0) { // Don't allow editing headers
      setEditingCell({ row, col });
      setEditValue(sheetData?.values[row - 1][col] || '');
    }
  };

  const handleCellBlur = () => {
    setEditingCell(null);
  };

  const handleCellKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      setEditingCell(null);
    }
  };

  const renderCellContent = (content: string, rowIndex?: number, colIndex?: number) => {
    if (!content) return '';

    // Get formatting for this cell if available
    const cellFormatting = rowIndex !== undefined && colIndex !== undefined
      ? sheetData?.formatting?.[rowIndex]?.values?.[colIndex]?.userEnteredFormat
      : null;

    if (cellFormatting) {
      const style: React.CSSProperties = {};
      
      // Apply text formatting
      if (cellFormatting.textFormat) {
        if (cellFormatting.textFormat.bold) style.fontWeight = 'bold';
        if (cellFormatting.textFormat.italic) style.fontStyle = 'italic';
        if (cellFormatting.textFormat.underline) style.textDecoration = 'underline';
        if (cellFormatting.textFormat.strikethrough) {
          style.textDecoration = style.textDecoration 
            ? `${style.textDecoration} line-through` 
            : 'line-through';
        }
        if (cellFormatting.textFormat.fontSize) {
          style.fontSize = `${cellFormatting.textFormat.fontSize}pt`;
        }
        if (cellFormatting.textFormat.foregroundColor) {
          const color = cellFormatting.textFormat.foregroundColor;
          style.color = `rgba(${Math.round(color.red * 255)}, ${Math.round(color.green * 255)}, ${Math.round(color.blue * 255)}, ${color.alpha || 1})`;
        }
      }

      // Apply background color
      if (cellFormatting.backgroundColor) {
        const bgColor = cellFormatting.backgroundColor;
        style.backgroundColor = `rgba(${bgColor.red * 255}, ${bgColor.green * 255}, ${bgColor.blue * 255}, ${bgColor.alpha || 1})`;
      }

      // Apply horizontal alignment
      if (cellFormatting.horizontalAlignment) {
        style.textAlign = cellFormatting.horizontalAlignment.toLowerCase() as 'left' | 'center' | 'right' | 'justify';
      }

      return <span className="rich-text-span" style={style}>{content}</span>;
    }

    // If no formatting or HTML content
    if (/<[^>]*>/.test(content)) {
      return <div dangerouslySetInnerHTML={{ __html: content }} />;
    }

    return content;
  };

  const handleDragOver = (e: React.DragEvent, row: number, col: number) => {
    e.preventDefault();
    e.stopPropagation();
    setDragOverCell({ row, col });
    e.dataTransfer.dropEffect = 'move';
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragOverCell(null);
  };

  const handleDrop = (e: React.DragEvent, row: number, col: number) => {
    e.preventDefault();
    e.stopPropagation();
    setDragOverCell(null);
    
    const fieldId = e.dataTransfer.getData('text/plain');
    if (fieldId) {
      const newMapping: CellMapping = {
        sourceId: fieldId,
        targetCell: { row, col }
      };
      setLastMapping(newMapping);
      setCellMappings((prev: CellMapping[]) => [...prev, newMapping]);
    }
  };

  // Update the SheetCell component in the render method
  const renderCell = (content: string, row: number, col: number) => {
    const mapping = cellMappings.find(
      mapping => mapping.targetCell.row === row && mapping.targetCell.col === col
    );

    const getFieldTitle = (sourceId: string) => {
      const fieldMap: { [key: string]: string } = {
        'process': 'Process & Impact',
        'components': 'Components',
        'assumptions': 'Assumptions',
        'hours': 'Hours',
        'notes': 'Notes',
        'outOfScope': 'Out of Scope Items',
        'roleHours': 'Role Hours',
        'hypercare': 'Hypercare'
      };
      return fieldMap[sourceId] || sourceId;
    };

    const handleDeleteMapping = (e: React.MouseEvent) => {
      e.stopPropagation();
      if (mapping) {
        setCellMappings(prev => prev.filter(m => 
          m.targetCell.row !== mapping.targetCell.row || 
          m.targetCell.col !== mapping.targetCell.col
        ));
      }
    };

    return (
      <SheetCell
        key={`cell-${row}-${col}`}
        isSelected={selectedCell?.row === row && selectedCell?.col === col}
        isHighlighted={dragOverCell?.row === row && dragOverCell?.col === col}
        isMapped={!!mapping}
        onClick={() => handleCellClick(row, col)}
        onDragOver={(e) => {
          if (!mapping) {
            handleDragOver(e, row, col);
          }
        }}
        onDragLeave={handleDragLeave}
        onDrop={(e) => {
          if (!mapping) {
            handleDrop(e, row, col);
          }
        }}
        width={columnWidths[col]}
        css={{
          cursor: mapping ? 'default' : 'default',
          '&[data-highlighted="true"]': {
            backgroundColor: 'var(--accent-secondary)',
            opacity: 0.7
          }
        }}
        data-highlighted={dragOverCell?.row === row && dragOverCell?.col === col}
        data-field-title={mapping ? getFieldTitle(mapping.sourceId) : undefined}
      >
        {mapping && (
          <button 
            className="delete-mapping"
            onClick={handleDeleteMapping}
            title="Remove mapping"
          >
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <line x1="18" y1="6" x2="6" y2="18"></line>
              <line x1="6" y1="6" x2="18" y2="18"></line>
            </svg>
          </button>
        )}
        {editingCell?.row === row && editingCell?.col === col ? (
          <CellInput
            value={editValue}
            onChange={(e) => setEditValue(e.target.value)}
            onBlur={handleCellBlur}
            onKeyDown={handleCellKeyDown}
            autoFocus
          />
        ) : (
          renderCellContent(content, row, col)
        )}
      </SheetCell>
    );
  };

  if (error) {
    return (
      <SheetPreview>
        <div css={{ color: 'var(--error-color)', padding: '1rem' }}>
          {error}
        </div>
      </SheetPreview>
    );
  }

  if (!sheetData) {
    return (
      <SheetPreview>
        <LoadingOverlay>Loading sheet data...</LoadingOverlay>
      </SheetPreview>
    );
  }

  return (
    <SheetPreview>
      <SheetGridContainer>
        <SheetGrid>
          {/* Corner and Column Headers */}
          <SheetRow height={sheetData?.rowMetadata?.[0]?.pixelSize ? `${sheetData.rowMetadata[0].pixelSize}px` : undefined}>
            <CornerCell />
            {sheetData?.headers.map((_, colIndex) => (
              <SheetCell
                key={`col-${colIndex}`}
                isHeader
                width={columnWidths[colIndex]}
              >
                {String.fromCharCode(65 + colIndex)}
              </SheetCell>
            ))}
          </SheetRow>

          {/* Header Row with Data */}
          <SheetRow height={sheetData?.rowMetadata?.[0]?.pixelSize ? `${sheetData.rowMetadata[0].pixelSize}px` : undefined}>
            <RowHeader>1</RowHeader>
            {sheetData?.headers.map((header, col) => renderCell(header, 0, col))}
          </SheetRow>

          {/* Data Rows */}
          {sheetData?.values.map((row, rowIndex) => (
            <SheetRow 
              key={`row-${rowIndex}`}
              height={sheetData?.rowMetadata?.[rowIndex + 1]?.pixelSize ? `${sheetData.rowMetadata[rowIndex + 1].pixelSize}px` : undefined}
            >
              <RowHeader>{rowIndex + 2}</RowHeader>
              {row.map((cell, colIndex) => renderCell(cell, rowIndex + 1, colIndex))}
            </SheetRow>
          ))}
        </SheetGrid>
      </SheetGridContainer>
      {loading && <LoadingOverlay>Loading...</LoadingOverlay>}
    </SheetPreview>
  );
};

// Add new interface for template state after other interfaces
interface TemplateState {
  spreadsheetId: string;
  cellMappings: CellMapping[];
}

// Add new interface for template data
interface ExcelTemplate {
  workbook: XLSX.WorkBook;
  filename: string;
}

// Add the ExcelIcon component after other icon components
const ExcelIcon = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
    <polyline points="14 2 14 8 20 8" />
    <path d="M8 13h8M8 17h8" />
  </svg>
);

function App() {
  const [isDarkMode, setIsDarkMode] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [rows, setRows] = useState<RowData[]>([]);
  const [selectedRows, setSelectedRows] = useState<number[]>([]);
  const [sheetsConfig, setSheetsConfig] = useState<GoogleSheetsConfig>({
    apiKey: 'AIzaSyCod0xneayl6ZA1JEaX-b8uQLzoHiPcuP4',
    spreadsheetId: '',
    range: 'Sheet1!A1:Z1000'
  });
  const [spreadsheetUrl, setSpreadsheetUrl] = useState('');
  const [urlError, setUrlError] = useState<string | null>(null);
  const [isDeleteModalOpen, setIsDeleteModalOpen] = useState(false);
  const [templateFile, setTemplateFile] = useState<TemplateFile | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [templateData, setTemplateData] = useState<TemplateData>({ 
    columns: [], 
    mapping: null,
    rows: []
  });
  const [saveMessage, setSaveMessage] = useState<{ text: string; success: boolean } | null>(null);
  const [selectedColumn, setSelectedColumn] = useState<string | null>(null);
  const [showMappingDialog, setShowMappingDialog] = useState(false);
  const [previewRows, setPreviewRows] = useState<any[]>([]);
  const [sheetUrl, setSheetUrl] = useState('');
  const [isValidUrl, setIsValidUrl] = useState(false);
  const [outOfScopeItems, setOutOfScopeItems] = useState<string[][]>([['']]);
  const [exporting, setExporting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  // Add state for selected out of scope items
  const [selectedOutOfScopeItems, setSelectedOutOfScopeItems] = useState<number[]>([]);
  const [isDeleteOutOfScopeModalOpen, setIsDeleteOutOfScopeModalOpen] = useState(false);
  // Add new interface for role hours
  const [roleHours, setRoleHours] = useState<RoleHours>({
    sa: '',
    consultant: '',
    pm: '',
    el: '',
    specialty: ''
  });
  // Add state for Hypercare after roleHours state
  const [hypercare, setHypercare] = useState<Hypercare>({
    hours: '',
    weeks: ''
  });
  // Add state for cell mappings after other state declarations
  const [cellMappings, setCellMappings] = useState<CellMapping[]>([]);
  const [lastMapping, setLastMapping] = useState<CellMapping | null>(null);
  const [templateState, setTemplateState] = useState<TemplateState>({
    spreadsheetId: '',
    cellMappings: []
  });
  // Add after other state declarations in App component
  const [excelTemplate, setExcelTemplate] = useState<ExcelTemplate | null>(null);

  const sensors = useSensors(
    useSensor(PointerSensor),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates,
    })
  );

  const addRow = () => {
    setRows([...rows, { ...emptyRow, id: crypto.randomUUID() }]);
  };

  const deleteSelectedRows = () => {
    const newRows = rows.filter((_, index) => !selectedRows.includes(index));
    setRows(newRows);
    setSelectedRows([]);
    setIsDeleteModalOpen(false);
  };

  const duplicateSelectedRows = () => {
    const newRows = [...rows];
    const duplicatedRows = selectedRows
      .sort((a, b) => b - a) // Sort in descending order to maintain correct indices
      .reduce((acc, index) => {
        acc.push({ ...rows[index] });
        return acc;
      }, [] as RowData[]);
    
    // Insert duplicated rows after the last selected row
    const lastSelectedIndex = Math.max(...selectedRows);
    newRows.splice(lastSelectedIndex + 1, 0, ...duplicatedRows);
    
    setRows(newRows);
    setSelectedRows([]);
  };

  const toggleRowSelection = (index: number) => {
    setSelectedRows(prev => {
      const isSelected = prev.includes(index);
      if (isSelected) {
        return prev.filter(i => i !== index);
      } else {
        return [...prev, index];
      }
    });
  };

  const handleSelectAll = (checked: boolean) => {
    if (checked) {
      setSelectedRows(rows.map((_, index) => index));
    } else {
      setSelectedRows([]);
    }
  };

  const updateRow = (index: number, field: keyof RowData, value: string) => {
    const newRows = [...rows];
    newRows[index] = { ...newRows[index], [field]: value };
    setRows(newRows);
  };

  const handleDragEnd = (event: DragEndEvent) => {
    const { active, over } = event;
    
    if (over && active.data.current?.type === 'mapping-item') {
      const [row, col] = (over.id as string).split('-').map(Number);
      const newMapping: CellMapping = {
        sourceId: active.id as string,
        targetCell: { row, col }
      };
      
      setLastMapping(newMapping);
      setCellMappings(prev => [...prev, newMapping]);
    }
  };

  const calculateTotal = () => {
    return rows.reduce((total, row) => {
      const hours = parseFloat(row.hours) || 0;
      return total + hours;
    }, 0).toFixed(1);
  };

  const handleExport = () => {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([
      ['Process and Impact', 'Components', 'Assumptions', 'Hours', 'Notes'], // Headers
      ...rows.map(row => [
        stripHtml(row.processAndImpact),
        stripHtml(row.components),
        stripHtml(row.assumptions),
        row.hours, // Hours is already a plain value
        stripHtml(row.notes)
      ])
    ]);

    XLSX.utils.book_append_sheet(workbook, worksheet, 'SOW');
    XLSX.writeFile(workbook, 'sow-data.xlsx');
  };

  const handleDragEnter = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  }, []);

  const validateFile = (file: File): boolean => {
    const validTypes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
      'application/vnd.ms-excel', // .xls
    ];

    if (!validTypes.includes(file.type)) {
      setUploadError('Please upload an Excel file (.xlsx or .xls)');
      return false;
    }

    if (file.size > 5 * 1024 * 1024) { // 5MB limit
      setUploadError('File size must be less than 5MB');
      return false;
    }

    return true;
  };

  const readExcelFile = async (file: File) => {
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data);
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      
      if (jsonData.length > 0) {
        const headers = jsonData[0] as string[];
        const rows = jsonData.slice(1); // Get all rows except headers
        
        setTemplateData({
          columns: headers,
          mapping: null,
          rows: rows
        });
        setPreviewRows(rows.slice(0, 5)); // Show first 5 rows in preview
      }
    } catch (error) {
      setUploadError('Error reading Excel file. Please try again.');
      console.error('Error reading Excel file:', error);
    }
  };

  const handleFile = useCallback((file: File) => {
    setUploadError(null);
    
    if (validateFile(file)) {
      setTemplateFile({
        name: file.name,
        size: file.size,
        type: file.type,
      });
      readExcelFile(file);
    }
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);

    const file = e.dataTransfer.files[0];
    if (file) {
      handleFile(file);
    }
  }, [handleFile]);

  const handleFileSelect = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      handleFile(file);
    }
  }, [handleFile]);

  const openFileDialog = () => {
    fileInputRef.current?.click();
  };

  const updateMapping = (field: keyof ColumnMapping, value: string) => {
    setTemplateData(prev => ({
      ...prev,
      mapping: {
        ...prev.mapping || {},
        [field]: value
      } as ColumnMapping
    }));
  };

  // Add new function to validate mapping
  const isMappingValid = () => {
    if (!templateData.mapping) return false;
    
    // Check if all required fields are mapped
    const requiredFields: (keyof ColumnMapping)[] = ['processAndImpact', 'components', 'assumptions', 'hours'];
    return requiredFields.every(field => templateData.mapping?.[field]);
  };

  // Add new function to save mapping
  const saveMapping = () => {
    if (!templateData.mapping) return;
    
    try {
      localStorage.setItem('templateMapping', JSON.stringify({
        fileName: templateFile?.name,
        mapping: templateData.mapping,
        lastUpdated: new Date().toISOString()
      }));
      
      setSaveMessage({
        text: 'Template mapping saved successfully!',
        success: true
      });
      
      // Clear success message after 3 seconds
      setTimeout(() => {
        setSaveMessage(null);
      }, 3000);
    } catch (error) {
      setSaveMessage({
        text: 'Error saving template mapping. Please try again.',
        success: false
      });
    }
  };

  // Add function to apply template data
  const applyTemplateData = () => {
    if (!templateData.mapping || !templateData.rows.length) return;

    const newRows = templateData.rows.map(row => {
      const columnToIndex = templateData.columns.reduce((acc, col, index) => {
        acc[col] = index;
        return acc;
      }, {} as { [key: string]: number });

      return {
        id: crypto.randomUUID(),
        processAndImpact: row[columnToIndex[templateData.mapping!.processAndImpact]] || '',
        components: row[columnToIndex[templateData.mapping!.components]] || '',
        assumptions: row[columnToIndex[templateData.mapping!.assumptions]] || '',
        hours: row[columnToIndex[templateData.mapping!.hours]]?.toString() || '',
        notes: row[columnToIndex[templateData.mapping!.notes]] || ''
      };
    });

    setRows(newRows);
    setShowSettings(false); // Return to SOW view
    setSaveMessage({
      text: 'Template data applied successfully!',
      success: true
    });

    // Clear success message after 3 seconds
    setTimeout(() => {
      setSaveMessage(null);
    }, 3000);
  };

  // Add function to handle column selection
  const handleColumnSelect = (column: string) => {
    setSelectedColumn(column);
    setShowMappingDialog(true);
  };

  // Add function to handle mapping selection
  const handleMappingSelect = (field: keyof ColumnMapping) => {
    if (selectedColumn) {
      updateMapping(field, selectedColumn);
      setShowMappingDialog(false);
      setSelectedColumn(null);
    }
  };

  // Add function to get mapped field for a column
  const getMappedField = (column: string): string | null => {
    if (!templateData.mapping) return null;
    
    for (const [field, mappedColumn] of Object.entries(templateData.mapping)) {
      if (mappedColumn === column) {
        return field;
      }
    }
    return null;
  };

  // Function to extract sheet ID from URL
  const extractSpreadsheetId = (url: string): string | null => {
    try {
      const regex = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/;
      const match = url.match(regex);
      return match ? match[1] : null;
    } catch (error) {
      return null;
    }
  };

  // Add new function to fetch and convert template
  const fetchAndStoreTemplate = async (spreadsheetId: string) => {
    try {
      setError(null);
      console.log('Fetching template for spreadsheet ID:', spreadsheetId);
      
      // Fetch the Google Sheet data using the sheets API
      const response = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Sheet1!A1:Z1000?key=${sheetsConfig.apiKey}`,
        {
          method: 'GET'
        }
      );

      console.log('Template fetch response status:', response.status);
      
      if (!response.ok) {
        const errorText = await response.text();
        console.error('Template fetch error response:', errorText);
        throw new Error(`Failed to fetch template data: ${response.status} ${response.statusText}`);
      }

      const data = await response.json();
      console.log('Template data received:', data.values ? `${data.values.length} rows` : 'No data');
      
      if (!data.values || data.values.length === 0) {
        throw new Error('Template appears to be empty. Please check that the sheet contains data.');
      }

      // Convert to XLSX format
      console.log('Converting to XLSX format...');
      const worksheet = XLSX.utils.aoa_to_sheet(data.values || []);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

      // Store template
      console.log('Storing template...');
      setExcelTemplate({
        workbook,
        filename: `template_${spreadsheetId}.xlsx`
      });

      return true;
    } catch (error) {
      console.error('Template fetch error:', error);
      const errorMessage = error instanceof Error 
        ? error.message 
        : 'Failed to fetch template. Please check your internet connection and try again.';
      setError(errorMessage);
      return false;
    }
  };

  // Update handleSpreadsheetUrlChange to include API key check
  const handleSpreadsheetUrlChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const url = e.target.value;
    setSpreadsheetUrl(url);
    
    if (url.trim() === '') {
      setUrlError(null);
      setSheetsConfig(prev => ({ ...prev, spreadsheetId: '' }));
      setExcelTemplate(null);
      return;
    }

    const spreadsheetId = extractSpreadsheetId(url);
    if (!spreadsheetId) {
      setUrlError('Please enter a valid Google Sheets URL');
      setSheetsConfig(prev => ({ ...prev, spreadsheetId: '' }));
      setExcelTemplate(null);
      return;
    }

    if (!sheetsConfig.apiKey) {
      setUrlError('Google Sheets API key is not configured. Please check your API key.');
      return;
    }

    setUrlError(null);
    setSheetsConfig(prev => ({ ...prev, spreadsheetId }));
    
    // Fetch and store template
    const success = await fetchAndStoreTemplate(spreadsheetId);
    if (!success) {
      setUrlError('Failed to fetch template. Please check the URL and ensure you have access to the sheet.');
    }
  };

  const handleSheetDataLoaded = (data: SheetData) => {
    console.log('Sheet data loaded:', data);
    // We'll implement mapping functionality here later
  };

  const handleOutOfScopeItemChange = (rowIndex: number, value: string) => {
    const newOutOfScopeItems = [...outOfScopeItems];
    newOutOfScopeItems[rowIndex] = [value];
    setOutOfScopeItems(newOutOfScopeItems);
  };

  const handleAddOutOfScopeItem = () => {
    setOutOfScopeItems([...outOfScopeItems, ['']]);
  };

  const handleDeleteOutOfScopeItem = (index: number) => {
    const newOutOfScopeItems = outOfScopeItems.filter((_, i) => i !== index);
    setOutOfScopeItems(newOutOfScopeItems);
  };

  const handleDuplicateOutOfScopeItem = (index: number) => {
    const newOutOfScopeItems = [...outOfScopeItems];
    newOutOfScopeItems.splice(index + 1, 0, [...outOfScopeItems[index]]);
    setOutOfScopeItems(newOutOfScopeItems);
  };

  // Update the exportToSheets function to include out of scope items
  const exportToSheets = async () => {
    try {
      setExporting(true);
      setError(null);

      if (!templateState.spreadsheetId || templateState.cellMappings.length === 0) {
        throw new Error('No template configuration found. Please set up template mappings in Settings first.');
      }

      // Create a new spreadsheet based on the template
      const response = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${templateState.spreadsheetId}/copy`,
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${sheetsConfig.apiKey}`,
            'Content-Type': 'application/json',
          }
        }
      );

      if (!response.ok) {
        throw new Error('Failed to create new spreadsheet from template');
      }

      const newSpreadsheet = await response.json();
      const newSpreadsheetId = newSpreadsheet.spreadsheetId;

      // Prepare the updates based on mappings
      const updates = templateState.cellMappings.map(mapping => {
        const { sourceId, targetCell } = mapping;
        let value = '';

        // Get the value based on the source field
        if (sourceId.startsWith('roleHours.')) {
          const role = sourceId.split('.')[1] as keyof RoleHours;
          value = roleHours[role];
        } else if (sourceId.startsWith('hypercare.')) {
          const field = sourceId.split('.')[1] as keyof Hypercare;
          value = hypercare[field];
        } else {
          switch (sourceId) {
            case 'process':
              value = rows.map(row => row.processAndImpact).join('\n\n');
              break;
            case 'components':
              value = rows.map(row => row.components).join('\n\n');
              break;
            case 'assumptions':
              value = rows.map(row => row.assumptions).join('\n\n');
              break;
            case 'hours':
              value = calculateTotal();
              break;
            case 'notes':
              value = rows.map(row => row.notes).join('\n\n');
              break;
            case 'outOfScope':
              value = outOfScopeItems.map(item => item[0]).join('\n\n');
              break;
          }
        }

        // Convert 0-based to A1 notation
        const column = String.fromCharCode(65 + targetCell.col);
        const row = targetCell.row + 1;
        const range = `Sheet1!${column}${row}`;

        return {
          range,
          values: [[value]],
          userEnteredFormat: {
            textFormat: {
              bold: false,  // Default format, will be overridden by rich text
            },
            wrapStrategy: 'WRAP'
          }
        };
      });

      // Batch update the new spreadsheet with values and formatting
      const updateResponse = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${newSpreadsheetId}:batchUpdate`,
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${sheetsConfig.apiKey}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            requests: [
              {
                updateCells: {
                  rows: updates.map(update => ({
                    values: [{
                      userEnteredValue: { stringValue: update.values[0][0] },
                      userEnteredFormat: update.userEnteredFormat
                    }]
                  })),
                  fields: 'userEnteredValue,userEnteredFormat',
                  range: {
                    sheetId: 0,
                    startRowIndex: Math.min(...templateState.cellMappings.map(m => m.targetCell.row)),
                    endRowIndex: Math.max(...templateState.cellMappings.map(m => m.targetCell.row)) + 1,
                    startColumnIndex: Math.min(...templateState.cellMappings.map(m => m.targetCell.col)),
                    endColumnIndex: Math.max(...templateState.cellMappings.map(m => m.targetCell.col)) + 1
                  }
                }
              }
            ]
          })
        }
      );

      if (!updateResponse.ok) {
        throw new Error('Failed to update spreadsheet with values');
      }

      // Open the new spreadsheet in a new tab
      window.open(`https://docs.google.com/spreadsheets/d/${newSpreadsheetId}`, '_blank');
      
    } catch (error) {
      console.error('Export error:', error);
      setError(error instanceof Error ? error.message : 'Failed to export to Sheets');
    } finally {
      setExporting(false);
    }
  };

  // Add handlers for out of scope items selection
  const toggleOutOfScopeItemSelection = (index: number) => {
    setSelectedOutOfScopeItems(prev => {
      const isSelected = prev.includes(index);
      if (isSelected) {
        return prev.filter(i => i !== index);
      } else {
        return [...prev, index];
      }
    });
  };

  const handleSelectAllOutOfScope = (checked: boolean) => {
    if (checked) {
      setSelectedOutOfScopeItems(outOfScopeItems.map((_, index) => index));
    } else {
      setSelectedOutOfScopeItems([]);
    }
  };

  const deleteSelectedOutOfScopeItems = () => {
    const newItems = outOfScopeItems.filter((_, index) => !selectedOutOfScopeItems.includes(index));
    setOutOfScopeItems(newItems);
    setSelectedOutOfScopeItems([]);
    setIsDeleteOutOfScopeModalOpen(false);
  };

  const duplicateSelectedOutOfScopeItems = () => {
    const newItems = [...outOfScopeItems];
    const duplicatedItems = selectedOutOfScopeItems
      .sort((a, b) => b - a)
      .reduce((acc, index) => {
        acc.push([...outOfScopeItems[index]]);
        return acc;
      }, [] as string[][]);
    
    const lastSelectedIndex = Math.max(...selectedOutOfScopeItems);
    newItems.splice(lastSelectedIndex + 1, 0, ...duplicatedItems);
    
    setOutOfScopeItems(newItems);
    setSelectedOutOfScopeItems([]);
  };

  // Add handler for updating role hours
  const handleRoleHoursChange = (role: keyof RoleHours, value: string) => {
    setRoleHours(prev => ({
      ...prev,
      [role]: value
    }));
  };

  // Add new styled component for the select dropdown after NumberInput
  const RoleSelect = styled.select`
    width: 100%;
    padding: 0.5rem;
    border: 1px solid var(--text-tertiary);
    border-radius: 0.25rem;
    background: var(--bg-primary);
    color: var(--text-primary);
    text-align: center;
    
    &:focus {
      outline: none;
      border-color: var(--accent-primary);
      box-shadow: 0 0 0 2px var(--accent-secondary);
    }
  `;

  // Add handler for Hypercare changes after handleRoleHoursChange
  const handleHypercareChange = (field: keyof Hypercare, value: string) => {
    setHypercare(prev => ({
      ...prev,
      [field]: value
    }));
  };

  const UndoButton = styled.button`
    background: var(--bg-secondary);
    color: var(--text-primary);
    border: 1px solid var(--text-tertiary);
    border-radius: 0.25rem;
    padding: 0.5rem 1rem;
    cursor: pointer;
    display: flex;
    align-items: center;
    gap: 0.5rem;
    font-size: 0.875rem;
    
    &:hover {
      background: var(--bg-tertiary);
    }
    
    &:disabled {
      opacity: 0.5;
      cursor: not-allowed;
    }
  `;

  // Add after the fetchAndStoreTemplate function
  const applyMappingsToTemplate = (): XLSX.WorkBook | null => {
    if (!excelTemplate || !templateState.cellMappings.length) {
      setError('No template or mappings found');
      return null;
    }

    try {
      // Create a deep copy of the template workbook
      const workbook = XLSX.utils.book_new();
      const originalSheet = excelTemplate.workbook.Sheets['Sheet1'];
      const newSheet = { ...originalSheet }; // Clone the sheet
      workbook.Sheets['Sheet1'] = newSheet;
      workbook.SheetNames = ['Sheet1'];

      // Apply each mapping
      templateState.cellMappings.forEach(mapping => {
        const { sourceId, targetCell } = mapping;
        let value = '';

        // Get the value based on the source field
        if (sourceId.startsWith('roleHours.')) {
          const role = sourceId.split('.')[1] as keyof RoleHours;
          value = roleHours[role];
        } else if (sourceId.startsWith('hypercare.')) {
          const field = sourceId.split('.')[1] as keyof Hypercare;
          value = hypercare[field];
        } else {
          switch (sourceId) {
            case 'process':
              value = rows.map(row => row.processAndImpact).join('\n\n');
              break;
            case 'components':
              value = rows.map(row => row.components).join('\n\n');
              break;
            case 'assumptions':
              value = rows.map(row => row.assumptions).join('\n\n');
              break;
            case 'hours':
              value = calculateTotal().toString();
              break;
            case 'notes':
              value = rows.map(row => row.notes).join('\n\n');
              break;
            case 'outOfScope':
              value = outOfScopeItems.map(item => item[0]).join('\n\n');
              break;
          }
        }

        // Convert to Excel cell reference (e.g., A1, B2)
        const cellRef = XLSX.utils.encode_cell({
          r: targetCell.row,
          c: targetCell.col
        });

        // Update cell in the sheet
        newSheet[cellRef] = {
          t: 's', // Type: string
          v: value, // Raw value
          w: value, // Formatted text
          s: { // Style - maintain wrapped text
            alignment: {
              wrapText: true,
              vertical: 'top'
            }
          }
        };
      });

      return workbook;
    } catch (error) {
      console.error('Error applying mappings:', error);
      setError('Failed to apply mappings to template');
      return null;
    }
  };

  // Update the export function with debugging
  const exportToExcel = async () => {
    console.log('Export started');
    console.log('Template state:', templateState);
    console.log('Excel template:', excelTemplate);
    
    try {
      setError(null);
      setExporting(true);

      // Check if we're in settings mode and have a template
      if (!excelTemplate?.workbook) {
        throw new Error('No template available. Please go to Settings, enter a valid Google Sheets URL, and wait for the template to load.');
      }

      if (!templateState.cellMappings.length) {
        throw new Error('No mappings configured. Please go to Settings and map your fields to template cells.');
      }

      console.log('Applying mappings...');
      // Apply mappings to get the final workbook
      const workbook = applyMappingsToTemplate();
      if (!workbook) {
        throw new Error('Failed to prepare workbook for export. Please check your mappings in Settings.');
      }

      console.log('Generating Excel file...');
      // Generate Excel file
      const excelBuffer = XLSX.write(workbook, {
        bookType: 'xlsx',
        type: 'array'
      });

      // Convert to Blob
      const blob = new Blob([excelBuffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });

      // Create download link
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      
      // Generate filename with timestamp
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
      link.download = `sow_export_${timestamp}.xlsx`;
      
      console.log('Triggering download...');
      // Trigger download
      document.body.appendChild(link);
      link.click();
      
      // Cleanup
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);
      console.log('Export completed successfully');

    } catch (error) {
      console.error('Export error:', error);
      setError(error instanceof Error ? error.message : 'Failed to export Excel file');
    } finally {
      setExporting(false);
    }
  };

  return (
    <GlobalStyle isDarkMode={isDarkMode}>
      <Container>
        <HeaderContainer>
          <Title>
            {showSettings ? 'Settings' : 'SOW Generator'}
          </Title>
          <HeaderActions>
            <NavButton
              onClick={() => setShowSettings(!showSettings)}
              title={showSettings ? 'Back to SOW Generator' : 'Settings'}
            >
              {showSettings ? <DocumentIcon /> : <SettingsIcon />}
            </NavButton>
            <ThemeToggleButton
              onClick={() => setIsDarkMode(!isDarkMode)}
              title={isDarkMode ? 'Switch to Light Mode' : 'Switch to Dark Mode'}
            >
              {isDarkMode ? <SunIcon /> : <MoonIcon />}
            </ThemeToggleButton>
          </HeaderActions>
        </HeaderContainer>

        {showSettings ? (
      <div>
            <ConfigSection>
              <h2 css={{ marginBottom: '1rem', color: 'var(--text-primary)' }}>Template Configuration</h2>
              <div css={{ 
                display: 'flex', 
                gap: '2rem',
                alignItems: 'flex-start',
                marginBottom: '1rem'  // Reduced from 2rem
              }}>
                <div css={{ flex: '1' }}>
                  <Label htmlFor="spreadsheetUrl">Google Sheets Template URL</Label>
                  <div css={{ display: 'flex', gap: '1rem', alignItems: 'flex-end' }}>
                    <Input
                      id="spreadsheetUrl"
                      value={spreadsheetUrl}
                      onChange={handleSpreadsheetUrlChange}
                      placeholder="Paste your Google Sheets URL here"
                      css={{ marginBottom: '0.5rem' }}
                    />
      </div>
                  {urlError && (
                    <div css={{ 
                      color: 'var(--error-color)',
                      fontSize: '0.875rem',
                      marginTop: '0.25rem'
                    }}>
                      {urlError}
      </div>
                  )}
                </div>
              </div>

              <div css={{ 
                fontSize: '0.875rem',
                color: 'var(--text-secondary)',
                marginTop: '0.5rem'
              }}>
                Steps to share your template:
                <ol css={{ marginTop: '0.5rem', paddingLeft: '1.25rem' }}>
                  <li>Open your Google Sheet template</li>
                  <li>Click "Share" in the top right</li>
                  <li>Set access to "Anyone with the link can view"</li>
                  <li>Copy the URL and paste it above</li>
                </ol>
              </div>
            </ConfigSection>

            {sheetsConfig.spreadsheetId && (
              <>
                <MappingContainer css={{ marginTop: '1rem' }}>
                  <div css={{
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    marginBottom: '1.5rem'
                  }}>
                    <MappingTitle css={{ margin: 0 }}>Map SOW Fields to Template</MappingTitle>
                    <div css={{ display: 'flex', gap: '1rem' }}>
                      <Button
                        onClick={() => {
                          setCellMappings([]);
                        }}
                        css={{
                          backgroundColor: 'var(--bg-secondary)',
                          color: 'var(--text-primary)',
                          border: '1px solid var(--text-tertiary)'
                        }}
                      >
                        Reset Mappings
                      </Button>
                      <Button
                        onClick={() => {
                          if (cellMappings.length > 0) {
                            setTemplateState({
                              spreadsheetId: sheetsConfig.spreadsheetId,
                              cellMappings: cellMappings
                            });
                            setShowSettings(false);
                          }
                        }}
                        disabled={cellMappings.length === 0}
                        css={{
                          backgroundColor: 'var(--accent-primary)',
                        }}
                      >
                        Apply Template
                      </Button>
                    </div>
                  </div>
                  <div css={{
                    display: 'grid',
                    gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
                    gap: '1rem',
                    marginBottom: '1.5rem'
                  }}>
                    {[
                      { id: 'process', label: 'Process & Impact' },
                      { id: 'components', label: 'Components' },
                      { id: 'assumptions', label: 'Assumptions' },
                      { id: 'hours', label: 'Hours' },
                      { id: 'notes', label: 'Notes' },
                      { id: 'outOfScope', label: 'Out of Scope Items' },
                      // Role Hours - Individual pills
                      { id: 'roleHours.sa', label: 'SA Hours' },
                      { id: 'roleHours.consultant', label: 'Consultant Hours' },
                      { id: 'roleHours.pm', label: 'PM Hours' },
                      { id: 'roleHours.el', label: 'EL Hours' },
                      { id: 'roleHours.specialty', label: 'Specialty Hours' },
                      // Hypercare components
                      { id: 'hypercare.hours', label: 'Hypercare Hours' },
                      { id: 'hypercare.weeks', label: 'Hypercare Weeks' }
                    ].map(field => {
                      const isMapped = cellMappings.some(mapping => mapping.sourceId === field.id);
                      return (
                        <div
                          key={field.id}
                          css={{
                            padding: '0.75rem',
                            background: 'var(--bg-tertiary)',
                            border: '2px solid var(--text-tertiary)',
                            borderRadius: '0.5rem',
                            cursor: isMapped ? 'not-allowed' : 'grab',
                            userSelect: 'none',
                            transition: 'all 0.2s ease',
                            opacity: isMapped ? 0.5 : 1,
                            filter: isMapped ? 'blur(1px)' : 'none',
                            pointerEvents: isMapped ? 'none' : 'all',
                            '&:hover': !isMapped ? {
                              borderColor: 'var(--accent-primary)',
                              transform: 'translateY(-1px)'
                            } : {}
                          }}
                          draggable={!isMapped}
                          onDragStart={(e) => {
                            if (!isMapped) {
                              e.dataTransfer.setData('text/plain', field.id);
                              e.dataTransfer.effectAllowed = 'move';
                            }
                          }}
                        >
                          {field.label}
                        </div>
                      );
                    })}
                  </div>
                  <div css={{
                    fontSize: '0.875rem',
                    color: 'var(--text-secondary)',
                    marginBottom: '1rem'
                  }}>
                    Drag and drop the fields above onto cells in the template below to create mappings.
                  </div>
                </MappingContainer>
                <GoogleSheetsPreview
                  config={sheetsConfig}
                  onDataLoaded={handleSheetDataLoaded}
                  setLastMapping={setLastMapping}
                  setCellMappings={setCellMappings}
                  cellMappings={cellMappings}
                />
              </>
            )}
          </div>
        ) : (
          // SOW Generator Content
          <>
            <ColumnsContainer>
              <HeaderRow>
                <ColumnHeader>
                  <HeaderCheckboxContainer>
                    <Checkbox
                      type="checkbox"
                      checked={selectedRows.length === rows.length && rows.length > 0}
                      onChange={(e) => handleSelectAll(e.target.checked)}
                    />
                  </HeaderCheckboxContainer>
                </ColumnHeader>
                <ColumnHeader>Process and Impact</ColumnHeader>
                <ColumnHeader>Components</ColumnHeader>
                <ColumnHeader>Assumptions</ColumnHeader>
                <ColumnHeader>Hours</ColumnHeader>
                <ColumnHeader>Notes</ColumnHeader>
              </HeaderRow>

              <DndContext
                sensors={sensors}
                collisionDetection={closestCenter}
                onDragEnd={handleDragEnd}
              >
                <SortableContext
                  items={rows.map(row => row.id)}
                  strategy={verticalListSortingStrategy}
                >
                  <RowsContainer>
                    {rows.map((row, index) => (
                      <SortableRow
                        key={row.id}
                        id={row.id}
                        index={index}
                        row={row}
                        isSelected={selectedRows.includes(index)}
                        onToggleSelect={() => toggleRowSelection(index)}
                        onUpdateRow={(field, value) => updateRow(index, field, value)}
                      />
                    ))}
                  </RowsContainer>
                </SortableContext>
              </DndContext>
            </ColumnsContainer>

            <TotalContainer>
              <ActionButtons>
                <AddRowButton onClick={addRow}>
                  <AddIcon />
                  Add Row
                </AddRowButton>
                <DeleteButton
                  onClick={() => setIsDeleteModalOpen(true)}
                  disabled={selectedRows.length === 0}
                >
                  <DeleteIcon />
                  Delete Row
                </DeleteButton>
                <DuplicateButton
                  onClick={duplicateSelectedRows}
                  disabled={selectedRows.length === 0}
                >
                  <DuplicateIcon />
                  Duplicate Row
                </DuplicateButton>
                <Button
                  onClick={() => {
                    console.log('Export button clicked');
                    exportToExcel();
                  }}
                  disabled={!templateState.cellMappings.length || exporting}
                  title={!templateState.cellMappings.length ? 'Please configure template mappings in Settings first' : ''}
                  css={{
                    backgroundColor: 'var(--accent-primary)',
                    opacity: (!templateState.cellMappings.length || exporting) ? 0.5 : 1,
                    cursor: (!templateState.cellMappings.length || exporting) ? 'not-allowed' : 'pointer',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '0.5rem',
                    '&:not(:disabled):hover': {
                      backgroundColor: 'var(--accent-secondary)'
                    }
                  }}
                >
                  <ExcelIcon />
                  {exporting ? 'Exporting...' : 'Export to Excel'}
                </Button>
              </ActionButtons>
              <TotalSection>
                <TotalLabel>Total Build Hours:</TotalLabel>
                <TotalValue>{calculateTotal()}</TotalValue>
              </TotalSection>
            </TotalContainer>

            {/* Add Out of Scope Items and Hours by Role sections in a flex container */}
            <div css={{ 
              display: 'flex', 
              gap: '2rem',
              marginTop: '2rem',
              alignItems: 'flex-start'  // This prevents containers from stretching to match heights
            }}>
              {/* Out of Scope Items section */}
              <ColumnsContainer css={{ 
                flex: 1,
                minWidth: '50%',
                maxWidth: '100%'
              }}>
                <HeaderRow css={{ 
                  marginBottom: '1rem',
                  display: 'flex',
                  justifyContent: 'space-between',
                  padding: '0.5rem',
                  width: '100%'
                }}>
                  <ColumnHeader css={{ 
                    display: 'flex', 
                    justifyContent: 'space-between', 
                    alignItems: 'center',
                    width: '100%',
                    padding: '0.5rem 1rem'
                  }}>
                    <div css={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
                      <HeaderCheckboxContainer>
                        <Checkbox
                          type="checkbox"
                          checked={selectedOutOfScopeItems.length === outOfScopeItems.length && outOfScopeItems.length > 0}
                          onChange={(e) => handleSelectAllOutOfScope(e.target.checked)}
                        />
                      </HeaderCheckboxContainer>
                      <span>Out of Scope Items</span>
                    </div>
                    <ActionButtons>
                      <AddRowButton 
                        onClick={handleAddOutOfScopeItem}
                        css={{ padding: '0.5rem 1rem', fontSize: '0.875rem' }}
                      >
                        <AddIcon />
                        Add Item
                      </AddRowButton>
                      <DeleteButton
                        onClick={() => setIsDeleteOutOfScopeModalOpen(true)}
                        disabled={selectedOutOfScopeItems.length === 0}
                        css={{ padding: '0.5rem 1rem', fontSize: '0.875rem' }}
                      >
                        <DeleteIcon />
                        Delete Item
                      </DeleteButton>
                      <DuplicateButton
                        onClick={duplicateSelectedOutOfScopeItems}
                        disabled={selectedOutOfScopeItems.length === 0}
                        css={{ padding: '0.5rem 1rem', fontSize: '0.875rem' }}
                      >
                        <DuplicateIcon />
                        Duplicate Item
                      </DuplicateButton>
                    </ActionButtons>
                  </ColumnHeader>
                </HeaderRow>

                <RowsContainer>
                  {outOfScopeItems.map((item, index) => (
                    <RowWrapper key={index} css={{ 
                      display: 'flex',
                      gap: '1rem',
                      width: '100%'
                    }}>
                      <CheckboxContainer css={{ flex: '0 0 40px' }}>
                        <Checkbox
                          type="checkbox"
                          checked={selectedOutOfScopeItems.includes(index)}
                          onChange={() => toggleOutOfScopeItemSelection(index)}
                        />
                      </CheckboxContainer>
                      <EditorContainer css={{ 
                        flex: '1',
                        '& .ProseMirror': { 
                          minHeight: '21px !important',  // Added !important
                          height: 'auto !important',     // Added !important
                          maxHeight: '200px !important', // Added !important
                          width: '100%',
                          maxWidth: '800px',
                          padding: '3px 8px',
                          overflowY: 'auto',
                          lineHeight: '1.2',  // Add line height control
                          fontSize: '14px'    // Control font size
                        },
                        '& .ProseMirror p': {  // Target paragraphs specifically
                          margin: '0',
                          padding: '0'
                        }
                      }}>
                        <RichTextEditor
                          content={item[0]}
                          onChange={(value) => handleOutOfScopeItemChange(index, value)}
                        />
                      </EditorContainer>
                    </RowWrapper>
                  ))}
                </RowsContainer>
              </ColumnsContainer>

              {/* Hours by Role section */}
              <ColumnsContainer css={{ 
                flex: '0 0 400px',  // Fixed width instead of flex: 1
                minWidth: 'auto',   // Remove minWidth constraint
                maxWidth: '400px'   // Match fixed width
              }}>
                <HeaderRow css={{ 
                  marginBottom: '1rem',
                  display: 'flex',
                  justifyContent: 'space-between',
                  padding: '0.5rem',
                  width: '100%'
                }}>
                  <ColumnHeader css={{ 
                    display: 'flex', 
                    justifyContent: 'space-between', 
                    alignItems: 'center',
                    width: '100%',
                    padding: '0.5rem 1rem'
                  }}>
                    <span>Hours by Role</span>
                  </ColumnHeader>
                </HeaderRow>

                <RowsContainer css={{
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '0.75rem',
                  padding: '0.5rem 1rem'
                }}>
                  {/* SA Hours */}
                  <div css={{
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    gap: '1rem'
                  }}>
                    <span css={{ 
                      color: 'var(--text-primary)',
                      fontWeight: 500,
                      flex: '1'
                    }}>
                      SA Hours Per Week
                    </span>
                    <NumberInputContainer css={{ width: '120px' }}>
                      <NumberInput
                        type="number"
                        min="0"
                        step="0.5"
                        value={roleHours.sa}
                        onChange={(e) => handleRoleHoursChange('sa', e.target.value)}
                      />
                    </NumberInputContainer>
                  </div>

                  {/* Consultant Hours */}
                  <div css={{
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    gap: '1rem'
                  }}>
                    <span css={{ 
                      color: 'var(--text-primary)',
                      fontWeight: 500,
                      flex: '1'
                    }}>
                      Consultant Hours Per Week
                    </span>
                    <NumberInputContainer css={{ width: '120px' }}>
                      <NumberInput
                        type="number"
                        min="0"
                        step="0.5"
                        value={roleHours.consultant}
                        onChange={(e) => handleRoleHoursChange('consultant', e.target.value)}
                      />
                    </NumberInputContainer>
                  </div>

                  {/* PM Hours */}
                  <div css={{
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    gap: '1rem'
                  }}>
                    <span css={{ 
                      color: 'var(--text-primary)',
                      fontWeight: 500,
                      flex: '1'
                    }}>
                      PM Hours Per Week
                    </span>
                    <NumberInputContainer css={{ width: '120px' }}>
                      <RoleSelect
                        value={roleHours.pm}
                        onChange={(e) => handleRoleHoursChange('pm', e.target.value)}
                      >
                        <option value="0">0</option>
                        <option value="20">20</option>
                        <option value="40">40</option>
                      </RoleSelect>
                    </NumberInputContainer>
                  </div>

                  {/* EL Hours */}
                  <div css={{
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    gap: '1rem'
                  }}>
                    <span css={{ 
                      color: 'var(--text-primary)',
                      fontWeight: 500,
                      flex: '1'
                    }}>
                      EL Hours Per Week
                    </span>
                    <NumberInputContainer css={{ width: '120px' }}>
                      <NumberInput
                        type="number"
                        min="0"
                        step="0.5"
                        value={roleHours.el}
                        onChange={(e) => handleRoleHoursChange('el', e.target.value)}
                      />
                    </NumberInputContainer>
                  </div>

                  {/* Specialty Resource Hours */}
                  <div css={{
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    gap: '1rem'
                  }}>
                    <span css={{ 
                      color: 'var(--text-primary)',
                      fontWeight: 500,
                      flex: '1'
                    }}>
                      Specialty Resource Hours
                    </span>
                    <NumberInputContainer css={{ width: '120px' }}>
                      <NumberInput
                        type="number"
                        min="0"
                        step="0.5"
                        value={roleHours.specialty}
                        onChange={(e) => handleRoleHoursChange('specialty', e.target.value)}
                      />
                    </NumberInputContainer>
                  </div>

                  {/* Add Hypercare section */}
                  <HeaderRow css={{ 
                    marginBottom: '1rem',
                    display: 'flex',
                    justifyContent: 'space-between',
                    padding: '0.5rem',
                    width: '100%',
                    marginTop: '1rem'  // Reduced from 2rem
                  }}>
                    <ColumnHeader css={{ 
                      display: 'flex', 
                      alignItems: 'center',  // Changed from space-between since we only have the title
                      width: '100%',
                      padding: '0.5rem 1rem',
                      justifyContent: 'flex-start'  // Add this to align text to the left
                    }}>
                      <span>Hypercare</span>
                    </ColumnHeader>
                  </HeaderRow>

                  <div css={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: '0.5rem',
                    padding: '0.5rem 1rem',
                    width: '100%'  // Added to ensure full width
                  }}>
                    <NumberInputContainer css={{ width: '80px' }}>
                      <NumberInput
                        type="number"
                        min="0"
                        step="0.5"
                        value={hypercare.hours}
                        onChange={(e) => handleHypercareChange('hours', e.target.value)}
                      />
                    </NumberInputContainer>
                    <span css={{ color: 'var(--text-primary)' }}>hours for</span>
                    <NumberInputContainer css={{ width: '80px' }}>
                      <NumberInput
                        type="number"
                        min="0"
                        step="1"
                        value={hypercare.weeks}
                        onChange={(e) => handleHypercareChange('weeks', e.target.value)}
                      />
                    </NumberInputContainer>
                    <span css={{ color: 'var(--text-primary)' }}>weeks</span>
                  </div>
                </RowsContainer>
              </ColumnsContainer>
            </div>

            {/* Add delete confirmation modal for out of scope items */}
            <ConfirmationModal
              isOpen={isDeleteOutOfScopeModalOpen}
              onClose={() => setIsDeleteOutOfScopeModalOpen(false)}
              onConfirm={deleteSelectedOutOfScopeItems}
              title="Delete Selected Items"
              message={`Are you sure you want to delete ${selectedOutOfScopeItems.length} selected item${selectedOutOfScopeItems.length === 1 ? '' : 's'}?`}
              confirmText="Delete"
              cancelText="Cancel"
            />
          </>
        )}
      </Container>
    </GlobalStyle>
  );
}

export default App;
