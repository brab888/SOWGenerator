import React from 'react';
import styled from '@emotion/styled';
import { useDraggable } from '@dnd-kit/core';

interface DraggableItemProps {
  id: string;
  label: string;
}

const ItemContainer = styled.div`
  margin: 0.5rem 0;
`;

const DraggablePill = styled.div<{ isDragging: boolean }>`
  background: var(--bg-tertiary);
  color: var(--text-primary);
  padding: 0.5rem 1rem;
  border-radius: 20px;
  display: inline-flex;
  align-items: center;
  cursor: grab;
  user-select: none;
  font-size: 0.875rem;
  border: 1px solid var(--text-tertiary);
  transition: all 0.2s ease;
  
  &:hover {
    background: var(--bg-secondary);
    border-color: var(--accent-primary);
  }

  ${({ isDragging }) => isDragging && `
    opacity: 0.5;
    cursor: grabbing;
    box-shadow: var(--shadow-md);
  `}
`;

const GroupContainer = styled.div`
  margin-bottom: 2rem;
`;

const GroupTitle = styled.h3`
  color: var(--text-primary);
  font-size: 1rem;
  margin-bottom: 1rem;
  padding-bottom: 0.5rem;
  border-bottom: 1px solid var(--text-tertiary);
`;

const DraggableItem: React.FC<DraggableItemProps> = ({ id, label }) => {
  const { attributes, listeners, setNodeRef, isDragging } = useDraggable({
    id: id,
    data: {
      type: 'mapping-item',
      label
    }
  });

  return (
    <ItemContainer>
      <DraggablePill
        ref={setNodeRef}
        isDragging={isDragging}
        {...listeners}
        {...attributes}
      >
        {label}
      </DraggablePill>
    </ItemContainer>
  );
};

const mainSectionItems = [
  { id: 'process-impact', label: 'Process and Impact' },
  { id: 'components', label: 'Components' },
  { id: 'assumptions', label: 'Assumptions' },
  { id: 'hours', label: 'Hours' },
  { id: 'notes', label: 'Notes' }
];

const outOfScopeItems = [
  { id: 'out-of-scope', label: 'Out of Scope Items' }
];

const singleCellItems = [
  { id: 'sa-hours', label: 'SA Hours Per Week' },
  { id: 'consultant-hours', label: 'Consultant Hours Per Week' },
  { id: 'pm-hours', label: 'PM Hours Per Week' },
  { id: 'el-hours', label: 'EL Hours Per Week' },
  { id: 'specialty-hours', label: 'Specialty Resource Hours' },
  { id: 'hypercare-hours', label: 'Hypercare Hours' },
  { id: 'hypercare-weeks', label: 'Hypercare Weeks' }
];

export const CellMappingPanel: React.FC = () => {
  return (
    <div css={{ padding: '1rem' }}>
      <GroupContainer>
        <GroupTitle>Main Sections</GroupTitle>
        {mainSectionItems.map((item) => (
          <DraggableItem key={item.id} {...item} />
        ))}
      </GroupContainer>

      <GroupContainer>
        <GroupTitle>Out of Scope</GroupTitle>
        {outOfScopeItems.map((item) => (
          <DraggableItem key={item.id} {...item} />
        ))}
      </GroupContainer>

      <GroupContainer>
        <GroupTitle>Individual Values</GroupTitle>
        {singleCellItems.map((item) => (
          <DraggableItem key={item.id} {...item} />
        ))}
      </GroupContainer>
    </div>
  );
};

export default CellMappingPanel; 