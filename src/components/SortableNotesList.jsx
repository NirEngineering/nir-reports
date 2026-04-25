import {
  DndContext,
  PointerSensor,
  TouchSensor,
  useSensor,
  useSensors,
  closestCenter,
} from '@dnd-kit/core';
import {
  SortableContext,
  useSortable,
  verticalListSortingStrategy,
  arrayMove,
} from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';

function NoteCard({ note, onEdit, onGenerate, onArchive, onDelete, isActive, dragHandleProps, isDragging }) {
  return (
    <div
      className="archive-item"
      style={{
        flexDirection: 'column',
        alignItems: 'stretch',
        gap: 6,
        opacity: isDragging ? 0.5 : 1,
        background: isDragging ? '#e0e7ff' : undefined,
      }}
    >
      {/* Title row with optional drag handle */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
        {isActive && (
          <div
            {...dragHandleProps}
            style={{
              color: '#94a3b8', fontSize: 20, cursor: 'grab',
              userSelect: 'none', touchAction: 'none', flexShrink: 0, lineHeight: 1,
            }}
            title="גרור לסידור מחדש"
          >⠿</div>
        )}
        <div style={{ fontWeight: 600, fontSize: 15, flex: 1, textAlign: 'right', direction: 'rtl' }}>
          {note.title || note.client || 'ללא כותרת'}
        </div>
      </div>

      {/* Sub-info */}
      <div style={{ textAlign: 'right', direction: 'rtl' }}>
        <div className="archive-item-sub">
          {note.client && <span>{note.client} • </span>}
          {note.created_at} • {note.updated_at}
        </div>
        {note.photos?.length > 0 && (
          <div className="archive-item-subject" style={{ color: '#64748b' }}>
            📷 {note.photos.length} תמונות
          </div>
        )}
      </div>

      {/* Single button row: ערוך | וורד | ארכיון | מחק */}
      <div style={{ display: 'flex', gap: 5, direction: 'rtl' }}>
        {isActive
          ? <button className="btn btn-sm btn-primary" onClick={() => onEdit(note)} style={{ fontSize: 12, flex: 1 }}>ערוך</button>
          : <div style={{ flex: 1 }} />
        }
        <button className="btn btn-sm btn-outline" onClick={() => onGenerate(note)} title="צור מסמך Word" style={{ fontSize: 12, flex: 1 }}>
          <span style={{ color: '#2B579A', fontWeight: 700 }}>W</span> Word
        </button>
        {isActive
          ? <button className="btn btn-sm btn-outline" onClick={() => onArchive(note.id, true)} style={{ fontSize: 12, flex: 1 }}>📁 ארכיון</button>
          : <button className="btn btn-sm btn-outline" onClick={() => onArchive(note.id, false)} style={{ fontSize: 12, flex: 1 }}>↩️ פעיל</button>
        }
        <button
          className="btn btn-sm btn-outline"
          style={{ color: '#ef4444', borderColor: '#ef4444', fontSize: 12, flex: 1 }}
          onClick={() => onDelete(note.id)}
        >🗑️ מחק</button>
      </div>
    </div>
  );
}

function SortableNoteCard(props) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } =
    useSortable({ id: props.note.id });

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
  };

  return (
    <div ref={setNodeRef} style={style}>
      <NoteCard {...props} dragHandleProps={{ ...attributes, ...listeners }} isDragging={isDragging} />
    </div>
  );
}

export default function SortableNotesList({ notes, onReorder, onEdit, onGenerate, onArchive, onDelete, isActive }) {
  const sensors = useSensors(
    useSensor(PointerSensor, { activationConstraint: { delay: 400, tolerance: 5 } }),
    useSensor(TouchSensor,   { activationConstraint: { delay: 400, tolerance: 5 } }),
  );

  const handleDragEnd = ({ active, over }) => {
    if (!over || active.id === over.id) return;
    const oldIdx = notes.findIndex(n => n.id === active.id);
    const newIdx = notes.findIndex(n => n.id === over.id);
    onReorder(arrayMove(notes, oldIdx, newIdx));
  };

  if (!isActive) {
    // Archived view — no drag, just render cards
    return (
      <div className="archive-list">
        {notes.map(note => (
          <NoteCard
            key={note.id}
            note={note}
            isActive={false}
            onEdit={onEdit}
            onGenerate={onGenerate}
            onArchive={onArchive}
            onDelete={onDelete}
            dragHandleProps={{}}
            isDragging={false}
          />
        ))}
      </div>
    );
  }

  return (
    <DndContext sensors={sensors} collisionDetection={closestCenter} onDragEnd={handleDragEnd}>
      <SortableContext items={notes.map(n => n.id)} strategy={verticalListSortingStrategy}>
        <div className="archive-list">
          {notes.map(note => (
            <SortableNoteCard
              key={note.id}
              note={note}
              isActive={true}
              onEdit={onEdit}
              onGenerate={onGenerate}
              onArchive={onArchive}
              onDelete={onDelete}
            />
          ))}
        </div>
      </SortableContext>
    </DndContext>
  );
}
