export type AssetItem = {
  name: string;
  value: number;
};

export type AppState = {
  last_filled_year?: number | string;
  last_filled_month?: number | string;
  leumi_balance?: number | string;
  cars?: AssetItem[];
  investments?: AssetItem[];
};

export type AssetsResponse = {
  cars: AssetItem[];
  investments: AssetItem[];
  leumi_balance: number | string | null;
};

export type CategoryGroup = {
  id: string;
  label: string;
  categories: string[];
};

export type ReviewItem = {
  name: string;
  cost: number;
  category: string;
  is_income?: boolean;
  /** Where the transaction came from, e.g. Max / לאומי עו"ש / לאומי כרטיס */
  source?: string;
  /** Transaction date as YYYY-MM-DD when available */
  date?: string | null;
  is_possible_duplicate: boolean;
  needs_manual_review?: boolean;
  error_reason?: string | null;
};

export type PrepareRunResponse = {
  run_token: string;
  review_token: string;
  source_excel: string;
  output_excel: string;
  allowed_categories: string[];
  category_groups?: CategoryGroup[];
  items: ReviewItem[];
  parse_errors: string[];
  warnings?: string[];
};

export type FinalizeRunResponse = {
  output_excel: string;
  yearly_output_excel?: string;
  drive_upload?: DriveUploadResult;
  state: AppState;
};

export type DriveStatusResponse = {
  configured: boolean;
  connected: boolean;
  email?: string | null;
  last_error?: string | null;
  updated_at?: string | null;
};

export type DriveConnectStartResponse = {
  authorization_url: string;
};

export type DriveUploadFileResult = {
  file_name?: string;
  file_id?: string;
  web_view_link?: string;
  action?: "created" | "updated";
  success: boolean;
};

export type DriveUploadResult = {
  attempted: boolean;
  connected: boolean;
  success: boolean;
  error?: string | null;
  folder_path?: string | null;
  monthly_file?: DriveUploadFileResult | null;
  yearly_file?: DriveUploadFileResult | null;
};

export type ProgressResponse = {
  step: string;
  updated_at: string | null;
  done: boolean;
};

export type ExpensesSummary = Record<string, number>;

export type TotalsTimelinePoint = {
  year: number;
  month: number;
  label: string;
  spending: number;
  income: number;
  investments: number;
};

export type TotalsTimelineResponse = {
  mode: "year" | "trailing" | "range";
  points: TotalsTimelinePoint[];
};

export type CategoriesResponse = {
  groups: CategoryGroup[];
  categories: string[];
};

export type CategoryTimelinePoint = {
  year: number;
  month: number;
  label: string;
  amount: number;
};

export type CategoryTimelineResponse = {
  mode: "year" | "trailing" | "range";
  category: string;
  points: CategoryTimelinePoint[];
};

export type PrepareRunPayload = {
  year: string;
  month: number;
  include_leumi: boolean;
  run_token: string;
};

export type FinalizeRunItem = {
  name: string;
  cost: number;
  category: string;
};

export type FinalizeRunPayload = {
  run_token: string;
  review_token: string;
  reviewed_items: FinalizeRunItem[];
};
