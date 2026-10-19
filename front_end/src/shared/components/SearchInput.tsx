import React from 'react';
import { Search } from 'lucide-react';

interface SearchInputProps {
  value: string;
  onChange: (value: string) => void;
}

export const SearchInput: React.FC<SearchInputProps> = React.memo(({ value, onChange }) => {
  return (
    <div className="w-full max-w-[400px]">
      <div className="e-input-group">
        <span className="e-input-group-icon e-icons flex items-center justify-center">
          <Search size={16} />
        </span>
        <input 
          className="e-input w-full" 
          type="text" 
          placeholder="Search mail..." 
          value={value}
          onChange={(e) => onChange(e.target.value)}
          autoComplete="off"
        />
      </div>
    </div>
  );
});

SearchInput.displayName = 'SearchInput';
