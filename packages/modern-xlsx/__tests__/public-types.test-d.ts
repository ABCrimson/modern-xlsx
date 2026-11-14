import { describe, expectTypeOf, it } from 'vitest';
import type {
  CalcPrData,
  ColorData,
  FileVersionData,
  MentionData,
  SharedItemsAttrs,
  SlicerPivotTable,
  WorkbookPrData,
} from '../src/index.js';
import type {
  CalcPrData as LiteCalcPrData,
  ColorData as LiteColorData,
  FileVersionData as LiteFileVersionData,
  MentionData as LiteMentionData,
  SharedItemsAttrs as LiteSharedItemsAttrs,
  SlicerPivotTable as LiteSlicerPivotTable,
  WorkbookPrData as LiteWorkbookPrData,
} from '../src/index-lite.js';
import type * as T from '../src/model/types.js';

describe('public type re-exports', () => {
  it('exports the seven additive wire types from index', () => {
    expectTypeOf<ColorData>().toEqualTypeOf<T.ColorData>();
    expectTypeOf<CalcPrData>().toEqualTypeOf<T.CalcPrData>();
    expectTypeOf<FileVersionData>().toEqualTypeOf<T.FileVersionData>();
    expectTypeOf<WorkbookPrData>().toEqualTypeOf<T.WorkbookPrData>();
    expectTypeOf<MentionData>().toEqualTypeOf<T.MentionData>();
    expectTypeOf<SharedItemsAttrs>().toEqualTypeOf<T.SharedItemsAttrs>();
    expectTypeOf<SlicerPivotTable>().toEqualTypeOf<T.SlicerPivotTable>();
  });

  it('exports the same seven from index-lite', () => {
    expectTypeOf<LiteColorData>().toEqualTypeOf<T.ColorData>();
    expectTypeOf<LiteCalcPrData>().toEqualTypeOf<T.CalcPrData>();
    expectTypeOf<LiteFileVersionData>().toEqualTypeOf<T.FileVersionData>();
    expectTypeOf<LiteWorkbookPrData>().toEqualTypeOf<T.WorkbookPrData>();
    expectTypeOf<LiteMentionData>().toEqualTypeOf<T.MentionData>();
    expectTypeOf<LiteSharedItemsAttrs>().toEqualTypeOf<T.SharedItemsAttrs>();
    expectTypeOf<LiteSlicerPivotTable>().toEqualTypeOf<T.SlicerPivotTable>();
  });
});
