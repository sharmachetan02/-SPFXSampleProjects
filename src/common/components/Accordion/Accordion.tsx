import * as React from 'react';
import styles from './Accordion.module.scss';
import SvgIcon from '../svgIcon/SvgIcon';
import *  as IconInterface from '../svgIcon/Interfaces';
import { Icon } from '@fluentui/react';

type OpenItems = string[]; // normalizing on arrays for simplicity

export interface AccordionProps {
    /** Allow multiple items open at once (v9: multiple) */
    multiple?: boolean;
    /** Allow closing the last open item (v9: collapsible) */
    collapsible?: boolean;
    /** Controlled open items (v9: openItems) */
    openItems?: OpenItems;
    /** Default open items (uncontrolled) (v9: defaultOpenItems) */
    defaultOpenItems?: OpenItems;
    /** Called when open state changes (v9: onOpenChange) */
    onOpenChange?: (event: React.SyntheticEvent, data: { openItems: OpenItems }) => void;
    /** Class name for the root */
    className?: string;
    /** Children: AccordionItem elements */
    children: React.ReactNode;
}

export interface AccordionItemProps {
    /** Unique value (v9: value) */
    value: string;
    /** Disabled state */
    disabled?: boolean;
    /** Children: AccordionHeader + AccordionPanel */
    children: React.ReactNode;
    className?: string;
}

export interface AccordionHeaderProps {
    /** Header content */
    children: React.ReactNode;
    /** Optional aria-label override */
    ariaLabel?: string;
    className?: string;
    /** Optional icon name for left icon */
    icon?: string; // NEW
}

export interface AccordionPanelProps {
    /** Whether to keep the panel mounted when collapsed */
    keepMounted?: boolean;
    className?: string;
    children: React.ReactNode;
}

/* =========================
   Context for state sharing
   ========================= */
interface AccordionCtx {
    registerHeader: (ref: HTMLButtonElement | null) => number;
    focusHeaderAt: (index: number) => void;
    getHeaderCount: () => number;

    isItemOpen: (value: string) => boolean;
    toggleItem: (value: string, ev: React.SyntheticEvent) => void;

    idPrefix: string;
    multiple: boolean;
    collapsible: boolean;
}

const AccordionContext = React.createContext<AccordionCtx | null>(null);
interface ItemCtx {
    value: string;
    disabled: boolean;
}
const ItemContext = React.createContext<ItemCtx | null>(null);


export const Accordion: React.FC<AccordionProps> = ({
    multiple = false,
    collapsible = false,
    openItems,
    defaultOpenItems,
    onOpenChange,
    className,
    children
}) => {

    const isControlled = Array.isArray(openItems);
    const [uncontrolledOpen, setUncontrolledOpen] = React.useState<OpenItems>(() => {
        if (Array.isArray(defaultOpenItems)) return Array.from(new Set(defaultOpenItems));
        return [];
    });
    const currentOpen: OpenItems = isControlled ? (openItems as OpenItems) : uncontrolledOpen;

    const headerRefs = React.useRef<HTMLButtonElement[]>([]);
    const registerHeader = (ref: HTMLButtonElement | null): number => {
        if (!ref) return -1;
        const idx = headerRefs.current.indexOf(ref);
        if (idx === -1) {
            headerRefs.current.push(ref);
            return headerRefs.current.length - 1;
        }
        return idx;
    };
    const focusHeaderAt = (index: number) => {
        const clamped = Math.max(0, Math.min(index, headerRefs.current.length - 1));
        const el = headerRefs.current[clamped];
        if (el) el.focus();
    };
    const getHeaderCount = () => headerRefs.current.length;

    const setOpen = (next: OpenItems, ev: React.SyntheticEvent) => {
        if (!isControlled) setUncontrolledOpen(next);
        onOpenChange?.(ev, { openItems: next });
    };

    const toggleItem = (value: string, ev: React.SyntheticEvent) => {
        const isOpen = currentOpen.includes(value);

        if (multiple) {
            // multiple open allowed
            if (isOpen) {
                const next = currentOpen.filter((v) => v !== value);
                if (!collapsible && next.length === 0) {
                    // keep at least one open
                    return;
                }
                setOpen(next, ev);
            } else {
                setOpen([...currentOpen, value], ev);
            }
        } else
            // single open
            if (isOpen) {
                // close it only if collapsible
                if (collapsible) setOpen([], ev);
            } else {
                setOpen([value], ev);
            }
    };

    const ctx: AccordionCtx = {
        registerHeader,
        focusHeaderAt,
        getHeaderCount,
        isItemOpen: (val) => currentOpen.includes(val),
        toggleItem,
        idPrefix: 'acc-' + Math.random().toString(36).slice(2),
        multiple,
        collapsible
    };

    return (
        <AccordionContext.Provider value={ctx}>
            <div className={[styles.accordionRoot, className].filter(Boolean).join(' ')} role="presentation">
                {children}
            </div>
        </AccordionContext.Provider>
    );
};

/* =========================
   Item
   ========================= */
export const AccordionItem: React.FC<AccordionItemProps> = ({ value, disabled = false, className, children }) =>
(
    <div className={[styles.item, className, disabled ? styles.disabled : ''].filter(Boolean).join(' ')} data-value={value} data-disabled={!!disabled}>
        <ItemContext.Provider value={{ value, disabled }}>
            {children}
        </ItemContext.Provider>
    </div>
);

/* =========================
   Header
   ========================= */
export const AccordionHeader: React.FC<AccordionHeaderProps> = ({ children, ariaLabel, className, icon }) => {
    const ctx = React.useContext(AccordionContext);
    const item = React.useContext(ItemContext);
    if (!ctx) throw new Error('AccordionHeader must be used within <Accordion>.');
    if (!item) throw new Error('AccordionHeader must be inside <AccordionItem>.');

    const { value, disabled } = item;
    const headerRef = React.useRef<HTMLButtonElement | null>(null);
    const [index, setIndex] = React.useState<number>(-1);

    React.useEffect(() => {
        if (headerRef.current) setIndex(ctx.registerHeader(headerRef.current));
    }, []);

    const open = ctx.isItemOpen(value);
    const contentId = `${ctx.idPrefix}-panel-${value}`;
    const headerId = `${ctx.idPrefix}-header-${value}`;

    const onClick = (ev: React.MouseEvent<HTMLButtonElement>) => {
        if (!disabled) ctx.toggleItem(value, ev);
    };

    const onKeyDown = (ev: React.KeyboardEvent<HTMLButtonElement>) => {
        if (ev.defaultPrevented) return;
        switch (ev.key) {
            case 'ArrowDown': ev.preventDefault(); ctx.focusHeaderAt(index + 1); break;
            case 'ArrowUp': ev.preventDefault(); ctx.focusHeaderAt(index - 1); break;
            case 'Home': ev.preventDefault(); ctx.focusHeaderAt(0); break;
            case 'End': ev.preventDefault(); ctx.focusHeaderAt(ctx.getHeaderCount() - 1); break;
            case 'Enter':
            case ' ': ev.preventDefault(); if (!disabled) ctx.toggleItem(value, ev); break;
        }
    };

    return (
        <div className={styles.headerRow}>
            <button
                ref={headerRef}
                id={headerId}
                type="button"
                className={[
                    styles.header,
                    className,
                    open ? styles.headerOpen : '',
                    disabled ? styles.headerDisabled : ''
                ].filter(Boolean).join(' ')}
                aria-expanded={open}
                aria-controls={contentId}
                aria-label={ariaLabel}
                onClick={onClick}
                onKeyDown={onKeyDown}
                disabled={disabled}
            >
                <span className={styles.headerContent}>
                    {icon && (
                        <span className={styles.headerIcon}>
                            <SvgIcon icon={IconInterface.Icon[icon]} />
                        </span>
                    )}
                    {children}
                </span>
                <Icon
                    iconName="ChevronDown"
                    className={[styles.chevron, open ? styles.chevronOpen : ''].join(' ')}
                    aria-hidden="true"
                />
            </button>
        </div>
    );
};

/* =========================
   Panel
   ========================= */
export const AccordionPanel: React.FC<AccordionPanelProps> = ({ keepMounted = true, className, children }) => {
    const ctx = React.useContext(AccordionContext);
    const item = React.useContext(ItemContext);
    if (!ctx) throw new Error('AccordionPanel must be used within <Accordion>.');
    if (!item) throw new Error('AccordionPanel must be inside <AccordionItem>.');

    const { value } = item;
    const open = ctx.isItemOpen(value);
    const contentId = `${ctx.idPrefix}-panel-${value}`;
    const headerId = `${ctx.idPrefix}-header-${value}`;

    // We keep it mounted and collapse via CSS for smooth animation and reliable state
    if (!keepMounted && !open) {
        // optional branch if you really want unmounted behavior
        return null;
    }

    return (
        <div
            id={contentId}
            role="region"
            aria-labelledby={headerId}
            className={[styles.panel, open ? styles.panelOpen : styles.panelClosed, className].filter(Boolean).join(' ')}
        >
            <div className={styles.panelInner}>
                {children}
            </div>
        </div>
    );
};
