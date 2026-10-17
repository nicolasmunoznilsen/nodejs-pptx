import { describe, expect, it } from 'vitest';
// Regression test for issue #31: EMU_PER_INCH and CORE_PROPS_REL must be public
// exports of the package entry point. Import ONLY from '../src/index.js' so this
// test fails if the re-export is removed, even though the constants still exist
// in '../src/ooxml/constants.js'.
import { CORE_PROPS_REL, EMU_PER_INCH } from '../src/index.js';

describe('public exports (issue #31)', () => {
  it('exposes EMU_PER_INCH with its documented value and type', () => {
    expect(EMU_PER_INCH).toBe(914400);
    expect(typeof EMU_PER_INCH).toBe('number');
  });

  it('exposes CORE_PROPS_REL with its documented value and type', () => {
    expect(CORE_PROPS_REL).toBe('http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties');
    expect(typeof CORE_PROPS_REL).toBe('string');
  });

  it('lists both names in the index module namespace keys', async () => {
    const namespace = await import('../src/index.js');
    const keys = Object.keys(namespace);
    expect(keys).toContain('EMU_PER_INCH');
    expect(keys).toContain('CORE_PROPS_REL');
  });

  it('does not leak unrelated relationship constants through the entry point', async () => {
    const namespace = await import('../src/index.js');
    const keys = Object.keys(namespace);
    for (const hidden of ['REL_NS', 'SLIDE_REL', 'MAIN_REL', 'SLIDE_MASTER_REL', 'SLIDE_LAYOUT_REL', 'THEME_REL', 'IMAGE_REL']) {
      expect(keys).not.toContain(hidden);
    }
  });
});
