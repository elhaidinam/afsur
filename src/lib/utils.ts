import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

export function safeNumeric(val: any): number {
  if (val === undefined || val === null || val === '') return NaN;
  if (typeof val === 'number') return val;
  const cleaned = String(val).replace(',', '.').replace(/[^\d.-]/g, '');
  const parsed = parseFloat(cleaned);
  return isNaN(parsed) ? 0 : parsed;
}

export function cleanName(name: string): string {
  return name.trim().replace(/\s+/g, ' ');
}
