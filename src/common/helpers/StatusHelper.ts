
export type StatusSemantic =
    | "success"   // Done
    | "warning"   // In Progress / Attention
    | "error"     // Major hurdles / Blocked
    | "info"      // Not Started / Neutral info
    | "neutral"   // Fallback

export enum LegalSvcProviderStatus {
    Active = 'Active',
    Inactive = 'Inactive',
    Revoked = 'Revoked',
}
export enum DPStatus {
    Active = 'Active',
    Inactive = 'Inactive',
    Revoked = 'Revoked',
}
export enum TPStatus {
    Active = 'Active',
    Inactive = 'Inactive',
    Revoked = 'Revoked',
}
export enum EutEntityStatus {
    Active = 'Active',
    Inactive = 'Inactive',
    Dissolved = 'Dissolved',
}
export enum FeeStatus {
    Active = 'Active',
    Inactive = 'Inactive',
    Revoked = 'Revoked',
}
export enum RAGStatus {
    MajorHurdles = 'Major hurdles',
    InProgress = 'In Progress',
    NotStarted = 'Not Started',
    Done = 'Done',
    Sanctioned = 'Sanctioned',
}
export enum PaymentStatus {
    Paid = 'Paid',
    PartiallyPaid = 'Partially Paid',
    Overpaid = 'Overpaid',
    Cancelled = 'Cancelled',
    Refunded = 'Refunded',
    Rejected = 'Rejected',
}
export enum SNPStatus {
    Commissioning = 'Commissioning',
    Active = 'Active',
    Maintenance = 'Maintenance',
    Decommissioned = 'Decommissioned',
}
export interface StatusViewProps {
    semantic: StatusSemantic;
    className: string;
    icon?: string;
}

const SUCCESS_SET: ReadonlySet<string> = new Set([
    LegalSvcProviderStatus.Active,
    DPStatus.Active,
    TPStatus.Active,
    EutEntityStatus.Active,
    FeeStatus.Active,
    RAGStatus.Done,
    PaymentStatus.Paid,
    SNPStatus.Active,
]);
const WARNING_SET: ReadonlySet<string> = new Set([
    LegalSvcProviderStatus.Inactive,
    DPStatus.Inactive,
    TPStatus.Active,
    EutEntityStatus.Inactive,
    FeeStatus.Inactive,
    RAGStatus.InProgress,
    PaymentStatus.Overpaid,
    PaymentStatus.PartiallyPaid,
    SNPStatus.Maintenance
]);
const ERROR_SET: ReadonlySet<string> = new Set([
    LegalSvcProviderStatus.Revoked,
    DPStatus.Revoked,
    TPStatus.Revoked,
    EutEntityStatus.Dissolved,
    FeeStatus.Revoked,
    RAGStatus.MajorHurdles,
    PaymentStatus.Rejected,
    PaymentStatus.Cancelled,
    SNPStatus.Decommissioned,
    RAGStatus.Sanctioned
]);
const INFO_SET: ReadonlySet<string> = new Set([
    RAGStatus.NotStarted,
    PaymentStatus.Refunded,
    SNPStatus.Commissioning
]);

export function getStatusViewProps(key): StatusViewProps {

    // Normalize synonyms you use in your project
    if (SUCCESS_SET.has(key)) {
        return toProps("success", "Completed");
    }
    if (WARNING_SET.has(key)) {
        return toProps("warning", "Clock");
    }
    if (ERROR_SET.has(key)) {
        return toProps("error", "ErrorCircle");
    }
    if (INFO_SET.has(key)) {
        return toProps("info", "Info");
    }
    return toProps("neutral", "CircleRing");
}
function toProps(semantic: StatusSemantic, icon: string): StatusViewProps {
    const classNameMap = {
        "success": "statusSuccess",
        "warning": "statusWarning",
        "error": "statusError",
        "info": "statusInfo",
        "neutral": "statusNeutral"
    };
    return {
        semantic,
        className: classNameMap[semantic],
        icon
    };
}
