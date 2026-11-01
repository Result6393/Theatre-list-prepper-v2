export declare const WORKER_URL: string;
export declare function deriveKeys(password: string): Promise<{ authToken: string; encKey: string }>;
export declare function authHashFor(password: string): Promise<string>;
export declare function isSignedIn(): boolean;
export declare function signIn(password: string): Promise<void>;
export declare function signOut(): void;
export declare function pullBeforeStart(
  app: string,
  localKey: string,
  timeoutMs?: number,
): Promise<'pulled' | 'pushed' | 'same' | 'offline' | 'signed-out'>;
export declare function pushSoon(app: string, localKey: string, delayMs?: number): void;
export declare function mountSyncControl(
  el: HTMLElement,
  app: string,
  localKey: string,
  onApplied?: () => void,
): () => void;
