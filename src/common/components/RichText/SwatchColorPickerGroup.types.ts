export interface ISwatchColor {
  color: string;
  id: string;
  label: string;
}
export interface ISwatchColorPickerGroupProps {
  groupColors: ISwatchColor[];
  groupText: string;
  onColorChanged: (string) => void;
  selectedColor?: string;
}

