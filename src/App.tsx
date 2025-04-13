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

const ExportIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
    <polyline points="7 10 12 15 17 10"></polyline>
    <line x1="12" y1="15" x2="12" y2="3"></line>
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
  values: string[][];
  headers: string[];
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
  padding: 1rem;
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
  border-radius: 4px;
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
}>`
  display: table-cell;
  padding: 3px 6px;
  border: 1px solid #e5e7eb;
  background: ${props => 
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
    background: ${props => props.isHeader ? '#e8eaed' : '#f8f9fa'};
  }

  p {
    margin: 0;
  }

  ul, ol {
    margin: 0;
    padding-left: 16px;
  }

  b, strong {
    font-weight: 600;
  }

  i, em {
    font-style: italic;
  }

  .rich-text-span {
    display: inline;
  }
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
}> = ({ config, onDataLoaded }) => {
  const [sheetData, setSheetData] = useState<SheetData | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [selectedCell, setSelectedCell] = useState<CellPosition | null>(null);
  const [editingCell, setEditingCell] = useState<CellPosition | null>(null);
  const [editValue, setEditValue] = useState('');
  const [columnWidths, setColumnWidths] = useState<string[]>([]);

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
      
      const sheetData = { 
        headers, 
        values,
        formatting: sheetMetadata.data?.[0]?.rowData || [],
        columnMetadata: sheetMetadata.data?.[0]?.columnMetadata || [],
        rowMetadata: sheetMetadata.data?.[0]?.rowMetadata || []
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
      <div css={{ display: 'flex', justifyContent: 'flex-end', marginBottom: '1rem' }}>
        <RefreshButton
          onClick={fetchSheetData}
          disabled={loading}
          title="Refresh template data"
        >
          <RefreshIcon />
        </RefreshButton>
      </div>
      <SheetGridContainer>
        <SheetGrid>
          {/* Corner and Column Headers */}
          <SheetRow height={sheetData?.rowMetadata?.[0]?.pixelSize ? `${sheetData.rowMetadata[0].pixelSize}px` : undefined}>
            <CornerCell />
            {sheetData.headers.map((_, colIndex) => (
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
            {sheetData.headers.map((header, col) => (
              <SheetCell
                key={`header-${col}`}
                width={columnWidths[col]}
              >
                {renderCellContent(header, 0, col)}
              </SheetCell>
            ))}
          </SheetRow>

          {/* Data Rows */}
          {sheetData.values.map((row, rowIndex) => (
            <SheetRow 
              key={`row-${rowIndex}`}
              height={sheetData?.rowMetadata?.[rowIndex + 1]?.pixelSize ? `${sheetData.rowMetadata[rowIndex + 1].pixelSize}px` : undefined}
            >
              <RowHeader>{rowIndex + 2}</RowHeader>
              {row.map((cell, colIndex) => (
                <SheetCell
                  key={`cell-${rowIndex}-${colIndex}`}
                  isSelected={selectedCell?.row === rowIndex + 1 && selectedCell?.col === colIndex}
                  onClick={() => handleCellClick(rowIndex + 1, colIndex)}
                  width={columnWidths[colIndex]}
                >
                  {editingCell?.row === rowIndex + 1 && editingCell?.col === colIndex ? (
                    <CellInput
                      value={editValue}
                      onChange={(e) => setEditValue(e.target.value)}
                      onBlur={handleCellBlur}
                      onKeyDown={handleCellKeyDown}
                      autoFocus
                    />
                  ) : (
                    renderCellContent(cell, rowIndex + 1, colIndex)
                  )}
                </SheetCell>
              ))}
            </SheetRow>
          ))}
        </SheetGrid>
      </SheetGridContainer>
      {loading && <LoadingOverlay>Loading...</LoadingOverlay>}
    </SheetPreview>
  );
};

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

    if (over && active.id !== over.id) {
      setRows((items) => {
        const oldIndex = items.findIndex((item) => item.id === active.id);
        const newIndex = items.findIndex((item) => item.id === over.id);

        // Update selected rows
        const newSelectedRows = new Set<number>();
        selectedRows.forEach(index => {
          if (index === oldIndex) {
            newSelectedRows.add(newIndex);
          } else if (index > oldIndex && index <= newIndex) {
            newSelectedRows.add(index - 1);
          } else if (index < oldIndex && index >= newIndex) {
            newSelectedRows.add(index + 1);
          } else {
            newSelectedRows.add(index);
          }
        });
        setSelectedRows(Array.from(newSelectedRows));

        return arrayMove(items, oldIndex, newIndex);
      });
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

  // Handle URL input change
  const handleSpreadsheetUrlChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const url = e.target.value;
    setSpreadsheetUrl(url);
    
    if (url.trim() === '') {
      setUrlError(null);
      setSheetsConfig(prev => ({ ...prev, spreadsheetId: '' }));
      return;
    }

    const spreadsheetId = extractSpreadsheetId(url);
    if (spreadsheetId) {
      setUrlError(null);
      setSheetsConfig(prev => ({ ...prev, spreadsheetId }));
    } else {
      setUrlError('Please enter a valid Google Sheets URL');
      setSheetsConfig(prev => ({ ...prev, spreadsheetId: '' }));
    }
  };

  const handleSheetDataLoaded = (data: SheetData) => {
    console.log('Sheet data loaded:', data);
    // We'll implement mapping functionality here later
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
              <div css={{ marginBottom: '2rem' }}>
                <Label htmlFor="spreadsheetUrl">Google Sheets Template URL</Label>
                <Input
                  id="spreadsheetUrl"
                  value={spreadsheetUrl}
                  onChange={handleSpreadsheetUrlChange}
                  placeholder="Paste your Google Sheets URL here"
                  css={{ marginBottom: '0.5rem' }}
                />
                {urlError && (
                  <div css={{ 
                    color: 'var(--error-color)',
                    fontSize: '0.875rem',
                    marginTop: '0.25rem'
                  }}>
                    {urlError}
      </div>
                )}
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
              </div>
            </ConfigSection>
            
            {sheetsConfig.spreadsheetId && (
              <GoogleSheetsPreview
                config={sheetsConfig}
                onDataLoaded={handleSheetDataLoaded}
              />
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
                <ExportButton onClick={handleExport}>
                  <ExportIcon />
                  Export to Sheets
                </ExportButton>
              </ActionButtons>
              <TotalSection>
                <TotalLabel>Total Build Hours:</TotalLabel>
                <TotalValue>{calculateTotal()}</TotalValue>
              </TotalSection>
            </TotalContainer>

            <ConfirmationModal
              isOpen={isDeleteModalOpen}
              onClose={() => setIsDeleteModalOpen(false)}
              onConfirm={deleteSelectedRows}
              title="Delete Selected Rows"
              message={`Are you sure you want to delete ${selectedRows.length} selected row${selectedRows.length === 1 ? '' : 's'}?`}
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
