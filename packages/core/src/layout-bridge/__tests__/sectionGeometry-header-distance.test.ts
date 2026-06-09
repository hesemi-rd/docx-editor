/**
 * Regression for #740 — `w:header="0"` (header pinned to the page top) was
 * treated as falsy and replaced with Word's 0.5in (48px) default, over-
 * reserving the header band and pushing content onto a second page. An explicit
 * 0 distance must be honored; only an ABSENT distance falls back to the default.
 */

import { describe, expect, test } from 'bun:test';
import { getMargins, DEFAULT_HF_DISTANCE_PX } from '../sectionGeometry';
import type { SectionProperties } from '../../types/document';

const base: SectionProperties = {
  marginTop: 813,
  marginRight: 1134,
  marginBottom: 1134,
  marginLeft: 1134,
};

describe('getMargins header/footer distance (#740)', () => {
  test('honors an explicit header/footer distance of 0 (not the default)', () => {
    const m = getMargins({ ...base, headerDistance: 0, footerDistance: 0 });
    expect(m.header).toBe(0);
    expect(m.footer).toBe(0);
  });

  test('falls back to the 0.5in default only when the distance is absent', () => {
    const m = getMargins(base);
    expect(m.header).toBe(DEFAULT_HF_DISTANCE_PX);
    expect(m.footer).toBe(DEFAULT_HF_DISTANCE_PX);
  });

  test('converts a non-zero distance from twips to px', () => {
    const m = getMargins({ ...base, headerDistance: 720, footerDistance: 720 });
    expect(m.header).toBe(48); // 720 twips = 0.5in = 48px
    expect(m.footer).toBe(48);
  });
});
