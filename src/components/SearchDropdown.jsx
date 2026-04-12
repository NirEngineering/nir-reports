import { useState, useRef, useEffect } from 'react';

export default function SearchDropdown({ value, onChange, options, placeholder, allowCustom = true }) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState(value || '');
  const ref = useRef(null);

  useEffect(() => {
    setSearch(value || '');
  }, [value]);

  useEffect(() => {
    const handler = (e) => {
      if (ref.current && !ref.current.contains(e.target)) setOpen(false);
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  const filtered = options.filter(o =>
    o.toLowerCase().includes(search.toLowerCase()) ||
    search === ''
  );

  const handleInput = (e) => {
    setSearch(e.target.value);
    setOpen(true);
    if (allowCustom) onChange(e.target.value);
  };

  const handleSelect = (item) => {
    setSearch(item);
    onChange(item);
    setOpen(false);
  };

  return (
    <div className="dropdown-wrapper" ref={ref}>
      <input
        type="text"
        className="form-input"
        value={search}
        onChange={handleInput}
        onFocus={() => setOpen(true)}
        placeholder={placeholder || 'חפש או הקלד...'}
        autoComplete="off"
      />
      {open && filtered.length > 0 && (
        <div className="dropdown-list">
          {filtered.map((item, i) => (
            <div
              key={i}
              className="dropdown-item"
              onMouseDown={() => handleSelect(item)}
            >
              {item}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
