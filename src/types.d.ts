declare module '*.svg' {
  const content: any;
  export default content;
}

interface RowData {
  id: string;
  processAndImpact: string;
  components: string;
  assumptions: string;
  hours: string;
  notes: string;
}

interface RichTextEditorProps {
  content: string;
  onChange: (content: string) => void;
}

interface ButtonProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  isDarkMode?: boolean;
}

interface ConfirmationModalProps {
  isOpen: boolean;
  onClose: () => void;
  onConfirm: () => void;
  title: string;
  message: string;
  confirmText?: string;
  cancelText?: string;
}

interface SortableRowProps {
  id: string;
  index: number;
  row: RowData;
  isSelected: boolean;
  onToggleSelect: () => void;
  onUpdateRow: (field: keyof RowData, value: string) => void;
} 