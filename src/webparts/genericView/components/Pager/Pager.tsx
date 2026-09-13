import * as React from 'react';
import { IconButton, DefaultButton, Dropdown, IDropdownOption } from '@fluentui/react';
import styles from '../View.module.scss';

export interface PagerProps {
  /** Total number of records */
  total: number;
  /** 0-based page index (controlled) */
  pageIndex: number;
  /** Page size (controlled) */
  pageSize: number;
  /** Change current page (0-based) */
  onPageChange: (page: number) => void;
  /** Change page size */
  onPageSizeChange: (size: number) => void;

  /** How many page buttons to show around current (default 5) */
  window?: number;
  /** Allowed page-size options (default [10,25,50,100]) */
  pageSizeOptions?: number[];
  /** Extra class on root */
  className?: string;
}

const Pager: React.FC<PagerProps> = ({
  total,
  pageIndex,
  pageSize,
  onPageChange,
  onPageSizeChange,
  window = 5,
  pageSizeOptions = [10, 25, 50, 100],
  className,
}) => {
  const pageCount = Math.max(1, Math.ceil(total / pageSize));
  const safeIndex = Math.min(Math.max(0, pageIndex), pageCount - 1);

  const canPrev = safeIndex > 0;
  const canNext = safeIndex < pageCount - 1;

  const start = Math.max(0, safeIndex - Math.floor(window / 2));
  const end = Math.min(pageCount - 1, start + window - 1);
  const pages = Array.from({ length: end - start + 1 }, (_, i) => start + i);

  const from = total === 0 ? 0 : safeIndex * pageSize + 1;
  const to = Math.min(total, (safeIndex + 1) * pageSize);

  const sizeOptions: IDropdownOption[] = pageSizeOptions.map((n) => ({ key: n, text: String(n) }));

  return (
    <div className={`${styles.pager} ${className ?? ''}`}>
      <div className={styles.centerGroup}>
        <span className={styles.rangeLabel}>{from}–{to} of {total}</span>

        <div className={styles.navGroup}>
          <IconButton
            iconProps={{ iconName: 'ChevronLeft' }}
            className={styles.navBtn}
            ariaLabel="Previous page"
            disabled={!canPrev}
            onClick={() => onPageChange(safeIndex - 1)}
          />

          <div className={styles.pageButtons}>
            {pages.map((p) => (
              <DefaultButton
                key={p}
                text={(p + 1).toString()}
                onClick={() => onPageChange(p)}
                className={`${styles.pageBtn} ${p === safeIndex ? styles.pageBtnActive : ''}`}
                aria-current={p === safeIndex ? 'page' : undefined}
              />
            ))}
          </div>

          <IconButton
            iconProps={{ iconName: 'ChevronRight' }}
            className={styles.navBtn}
            ariaLabel="Next page"
            disabled={!canNext}
            onClick={() => onPageChange(safeIndex + 1)}
          />
        </div>
      </div>

      <div className={styles.rowsGroup}>
        <span className={styles.rowsLabel}>Rows:</span>
        <Dropdown
          className={styles.rowsDropdown}
          selectedKey={pageSize}
          options={sizeOptions}
          onChange={(_, opt) => {
            const next = (opt?.key as number) ?? pageSize;
            if (next && next !== pageSize) onPageSizeChange(next);
          }}
          ariaLabel="Rows per page"
        />
      </div>
    </div>
  );
};

export default Pager;
