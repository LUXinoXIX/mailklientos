// Typy TLS zdieľané s rendererom (bez importu node:tls).

export type Protocol = 'imap' | 'smtp';

export interface CertInfo {
  subject: string;
  issuer: string;
  validFrom: string;
  validTo: string;
  fingerprint256: string;
  selfSigned: boolean;
  altNames?: string;
}

export interface ProbeResult {
  ok: boolean;
  /** Spojenie sa podarilo, ale certifikát nie je dôveryhodný podľa politiky konta. */
  needsTrust?: boolean;
  protocol?: string | null;
  cipher?: string;
  cert?: CertInfo;
  /** Kód chyby overenia certifikátu (napr. DEPTH_ZERO_SELF_SIGNED_CERT). */
  verifyError?: string;
  hostnameMismatch?: boolean;
  /** Text pre používateľa (slovensky). */
  message: string;
  /** Odporúčaná oprava (slovensky). */
  hint?: string;
  /** Odporúča sa zapnúť režim pre staré servery. */
  suggestLegacy?: boolean;
}
