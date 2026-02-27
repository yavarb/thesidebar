/** Shared track-changes state */
let trackChangesEnabled = false;

export function setTrackChanges(enabled: boolean): void {
  trackChangesEnabled = enabled;
}

export function isTrackChangesEnabled(): boolean {
  return trackChangesEnabled;
}
