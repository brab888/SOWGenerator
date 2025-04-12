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
    box-shadow: 0 0 0 2px var(--accent-secondary);
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

function App() {
  const [isDarkMode, setIsDarkMode] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [rows, setRows] = useState<RowData[]>([{ ...emptyRow }]);
  const [selectedRows, setSelectedRows] = useState<Set<number>>(new Set());
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
    const newRows = rows.filter((_, index) => !selectedRows.has(index));
    setRows(newRows);
    setSelectedRows(new Set());
    setIsDeleteModalOpen(false);
  };

  const duplicateSelectedRows = () => {
    const newRows = [...rows];
    const rowsToDuplicate = Array.from(selectedRows).sort((a, b) => b - a);
    
    rowsToDuplicate.forEach(index => {
      const duplicatedRow = {
        ...rows[index],
        id: crypto.randomUUID()
      };
      newRows.splice(index + 1, 0, duplicatedRow);
    });

    setRows(newRows);
    setSelectedRows(new Set());
  };

  const toggleRowSelection = (index: number) => {
    const newSelected = new Set(selectedRows);
    if (newSelected.has(index)) {
      newSelected.delete(index);
    } else {
      newSelected.add(index);
    }
    setSelectedRows(newSelected);
  };

  const handleSelectAll = (checked: boolean) => {
    if (checked) {
      const allIndices = new Set(rows.map((_, index) => index));
      setSelectedRows(allIndices);
    } else {
      setSelectedRows(new Set());
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
        setSelectedRows(newSelectedRows);

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
  const extractSheetId = (url: string) => {
    try {
      const parsedUrl = new URL(url);
      if (!parsedUrl.hostname.includes('docs.google.com')) {
        return null;
      }
      
      // Extract the ID from various Google Sheets URL formats
      const matches = url.match(/\/d\/(.*?)([\/\?]|$)/);
      return matches ? matches[1] : null;
    } catch {
      return null;
    }
  };

  // Handle URL input change
  const handleUrlChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const url = e.target.value;
    setSheetUrl(url);
    
    const sheetId = extractSheetId(url);
    setIsValidUrl(!!sheetId);
    
    if (sheetId) {
      // Clear any previous error
      setUploadError(null);
    } else if (url) {
      setUploadError('Please enter a valid Google Sheets URL');
    }
  };

  return (
    <GlobalStyle isDarkMode={isDarkMode}>
      <Container>
        <HeaderContainer>
          <div style={{ width: '88px' }} /> {/* Spacer to balance the header */}
          <Title>{showSettings ? 'Settings' : 'SOW Generator'}</Title>
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
          <>
            <UploadContainer>
              <h2 style={{ 
                color: 'var(--text-primary)', 
                marginTop: 0, 
                marginBottom: '1rem' 
              }}>
                Import Template
              </h2>
              
              <Instructions>
                1. Open your Google Sheet<br />
                2. Click "Share" and set to "Anyone with the link can view"<br />
                3. Copy the URL and paste it below
              </Instructions>

              <URLInput
                type="url"
                placeholder="Paste your Google Sheets URL here"
                value={sheetUrl}
                onChange={handleUrlChange}
              />

              {uploadError && (
                <ErrorMessage>{uploadError}</ErrorMessage>
              )}

              {isValidUrl && (
                <Button
                  onClick={() => {/* We'll implement this next */}}
                  style={{ 
                    marginTop: '1rem',
                    background: 'var(--accent-primary)'
                  }}
                >
                  Load Template
                </Button>
              )}
            </UploadContainer>
            
            {templateData.columns.length > 0 && (
              <>
                <PreviewContainer>
                  <PreviewTable>
                    <thead>
                      <tr>
                        {templateData.columns.map((column) => {
                          const mappedField = getMappedField(column);
                          return (
                            <PreviewHeader
                              key={column}
                              isSelected={!!mappedField}
                              onClick={() => handleColumnSelect(column)}
                              title={mappedField ? `Mapped to ${mappedField}` : 'Click to map this column'}
                            >
                              {column}
                              {mappedField && (
                                <div style={{ fontSize: '0.75rem', marginTop: '0.25rem' }}>
                                  ↓ {mappedField}
                                </div>
                              )}
                            </PreviewHeader>
                          );
                        })}
                      </tr>
                    </thead>
                    <tbody>
                      {previewRows.map((row, rowIndex) => (
                        <tr key={rowIndex}>
                          {templateData.columns.map((column, colIndex) => (
                            <PreviewCell key={`${rowIndex}-${colIndex}`}>
                              {row[colIndex]}
                            </PreviewCell>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </PreviewTable>
                </PreviewContainer>

                <SaveButton
                  onClick={saveMapping}
                  disabled={!isMappingValid()}
                  style={{ margin: '1.5rem auto', display: 'block' }}
                >
                  Save Template Mapping
                </SaveButton>

                {templateData.rows.length > 0 && templateData.mapping && (
                  <ApplyButton
                    onClick={applyTemplateData}
                    disabled={!isMappingValid()}
                    style={{ margin: '1rem auto', display: 'block' }}
                  >
                    Apply Template Data ({templateData.rows.length} rows)
                  </ApplyButton>
                )}

                {saveMessage && (
                  <SaveMessage success={saveMessage.success}>
                    {saveMessage.success ? (
                      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" width="16" height="16">
                        <path d="M20 6L9 17l-5-5" stroke="currentColor" strokeWidth="2" fill="none"/>
                      </svg>
                    ) : (
                      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" width="16" height="16">
                        <path d="M12 9v4M12 17h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" stroke="currentColor" strokeWidth="2" fill="none"/>
                      </svg>
                    )}
                    {saveMessage.text}
                  </SaveMessage>
                )}
              </>
            )}

            {showMappingDialog && selectedColumn && (
              <MappingOverlay onClick={() => setShowMappingDialog(false)}>
                <MappingDialog onClick={(e) => e.stopPropagation()}>
                  <h3 style={{ margin: '0 0 1rem 0', color: 'var(--text-primary)' }}>
                    Map column: {selectedColumn}
                  </h3>
                  <MappingOptions>
                    <MappingOption
                      isSelected={templateData.mapping?.processAndImpact === selectedColumn}
                      onClick={() => handleMappingSelect('processAndImpact')}
                    >
                      Process and Impact
                    </MappingOption>
                    <MappingOption
                      isSelected={templateData.mapping?.components === selectedColumn}
                      onClick={() => handleMappingSelect('components')}
                    >
                      Components
                    </MappingOption>
                    <MappingOption
                      isSelected={templateData.mapping?.assumptions === selectedColumn}
                      onClick={() => handleMappingSelect('assumptions')}
                    >
                      Assumptions
                    </MappingOption>
                    <MappingOption
                      isSelected={templateData.mapping?.hours === selectedColumn}
                      onClick={() => handleMappingSelect('hours')}
                    >
                      Hours
                    </MappingOption>
                    <MappingOption
                      isSelected={templateData.mapping?.notes === selectedColumn}
                      onClick={() => handleMappingSelect('notes')}
                    >
                      Notes
                    </MappingOption>
                  </MappingOptions>
                  <Button 
                    onClick={() => setShowMappingDialog(false)}
                    style={{ width: '100%', marginTop: '1rem' }}
                  >
                    Cancel
                  </Button>
                </MappingDialog>
              </MappingOverlay>
            )}
          </>
        ) : (
          // SOW Generator Content
          <>
            <ColumnsContainer>
              <HeaderRow>
                <ColumnHeader>
                  <HeaderCheckboxContainer>
                    <Checkbox
                      type="checkbox"
                      checked={selectedRows.size === rows.length && rows.length > 0}
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
                        isSelected={selectedRows.has(index)}
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
                  disabled={selectedRows.size === 0}
                >
                  <DeleteIcon />
                  Delete Row
                </DeleteButton>
                <DuplicateButton
                  onClick={duplicateSelectedRows}
                  disabled={selectedRows.size === 0}
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
              message={`Are you sure you want to delete ${selectedRows.size} selected row${selectedRows.size === 1 ? '' : 's'}?`}
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
