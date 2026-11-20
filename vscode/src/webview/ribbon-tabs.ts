// Ribbon tabs: which groups each tab contains, in which order. All commands go through
// `api.cmd(name, …)` — the names are implemented once in `main.ts` (`ribbonCommands`).
import { h } from './util';
import { t, data, withAccel, kbd } from './i18n';
import { MenuItem } from './menu';
import {
  THEMES,
  LAYOUTS,
  TRANSITIONS,
  ANIM_CLASSES,
  ANIM_TYPES,
  ANIM_DIRECTIONS,
  FONT_LIST,
  FONT_SIZES
} from '@fefeding/ppt-parser';
import type { RibbonApi, RibbonState } from './ribbon';

export type RibbonTab = {
  id: string;
  label: string;
  panel: HTMLElement;
  /** Only meaningful in edit mode (hidden from the tab strip while previewing). */
  editOnly: boolean;
};

type Widget = { el: HTMLElement };

export type RibbonBuilders = {
  bigButton(o: {
    ico: string;
    label: string;
    title?: string;
    accel?: string;
    needs?: string | string[];
    onClick: () => void;
  }): Widget;
  miniButton(o: {
    ico?: string;
    glyph?: string;
    glyphStyle?: string;
    title: string;
    needs?: string | string[];
    active?: (s: RibbonState) => boolean;
    onClick: () => void;
  }): Widget;
  splitButton(o: {
    ico: string;
    label: string;
    accel?: string;
    needs?: string | string[];
    onClick: () => void;
    menu: () => MenuItem[];
  }): Widget;
  selectBox(o: {
    label: string;
    options: { value: string; label: string }[];
    width?: number;
    needs?: string | string[];
    onChange: (v: string) => void;
    onSync?: (s: RibbonState) => string;
  }): Widget & { setOptions(options: { value: string; label: string }[]): void };
  colorBox(o: {
    title: string;
    ico: string;
    needs?: string | string[];
    onInput: (v: string) => void;
    onSync?: (s: RibbonState) => string;
  }): Widget;
  gallery<T>(
    items: T[],
    renderItem: (item: T) => { el: HTMLElement; isActive?: (s: RibbonState) => boolean },
    opts?: { needs?: string | string[]; caption?: string }
  ): Widget;
  chip(
    label: string,
    title: string,
    onClick: () => void,
    active?: (s: RibbonState) => boolean
  ): { el: HTMLElement; isActive?: (s: RibbonState) => boolean };
  group(title: string, rows: (Widget | null)[][], editOnly?: boolean): HTMLElement;
  advanceCheck(): Widget;
  animListBox(): Widget;
};

const DEFAULT_TRANS_SPEED = '800';

/** Animation presets offered by the Animations tab. */
const ANIM_PRESETS: { cls: string; type: string; label: string }[] = [
  { cls: 'entr', type: 'fadeIn', label: 'data.animType.fadeIn' },
  { cls: 'entr', type: 'flyIn', label: 'data.animType.flyIn' },
  { cls: 'entr', type: 'wipeIn', label: 'data.animType.wipeIn' },
  { cls: 'entr', type: 'zoomIn', label: 'data.animType.zoomIn' },
  { cls: 'entr', type: 'riseUp', label: 'data.animType.riseUp' },
  { cls: 'emph', type: 'pulse', label: 'data.animType.pulse' }
];

const LINE_SPACING = ['0.9', '1', '1.15', '1.5', '2'];
// First entry must match the draft below, otherwise the select shows a value that isn't used.
const ANIM_DURATIONS = [500, 300, 800, 1200];

export function buildTabs(b: RibbonBuilders, api: RibbonApi): RibbonTab[] {
  // ---------- shared option lists ----------
  const layoutOptions = LAYOUTS.map((l: any) => ({ value: l.id, label: data('layout', l.id, l.name) }));
  const layoutMenu = (run: (id: string) => void): MenuItem[] => [
    { type: 'label', label: t('cmd.layout') },
    ...layoutOptions.map((o) => ({ label: o.label, icon: 'layout', onClick: () => run(o.value) }))
  ];
  const fontOptions = FONT_LIST.map((f: string) => ({ value: f, label: data('font', f, f) }));
  const sizeOptions = FONT_SIZES.map((n: number) => ({ value: String(n), label: String(n) }));
  const speedOptions = [
    { value: '1500', label: data('speed', 'slow', 'Slow') },
    { value: DEFAULT_TRANS_SPEED, label: data('speed', 'normal', 'Normal') },
    { value: '400', label: data('speed', 'fast', 'Fast') }
  ];
  const triggerOptions = ['onClick', 'withPrev', 'afterPrev'].map((v) => ({ value: v, label: data('trigger', v, v) }));
  const directionOptions = ANIM_DIRECTIONS.map((d: any) => ({ value: d.value, label: data('animDir', d.value, d.name) }));
  const classOptions = ANIM_CLASSES.map((c: any) => ({ value: c.value, label: data('animClass', c.value, c.name) }));
  const typeOptionsFor = (cls: string) => (ANIM_TYPES[cls] || []).map((x: any) => ({ value: x.value, label: data('animType', x.value, x.name) }));
  const sizeKeyOptions = ['16:9', '16:10', '4:3'].map((k) => ({ value: k, label: data('size', k, k) }));

  // Animation draft: the Animations tab is a form, not a live view of the store.
  // 默认 afterPrev（自动播放），与属性面板一致：点效果后 F5 就能立刻看到，不必先点一下。
  const anim = { cls: 'entr', type: 'fadeIn', trigger: 'afterPrev', direction: 'l', duration: 500 };
  const addAnim = (over: Partial<typeof anim>) => {
    Object.assign(anim, over);
    api.cmd('addAnim', { ...anim });
  };

  const textAlignRow = (): Widget[] => [
    b.miniButton({ ico: 'alignLeft', title: t('cmd.alignLeft'), needs: ['edit', 'text'], active: (s) => s.caps.align === 'left', onClick: () => api.cmd('setAlign', 'left') }),
    b.miniButton({ ico: 'alignCenter', title: t('cmd.alignCenter'), needs: ['edit', 'text'], active: (s) => s.caps.align === 'center', onClick: () => api.cmd('setAlign', 'center') }),
    b.miniButton({ ico: 'alignRight', title: t('cmd.alignRight'), needs: ['edit', 'text'], active: (s) => s.caps.align === 'right', onClick: () => api.cmd('setAlign', 'right') }),
    b.miniButton({ ico: 'alignJustify', title: t('cmd.alignJustify'), needs: ['edit', 'text'], active: (s) => s.caps.align === 'justify', onClick: () => api.cmd('setAlign', 'justify') }),
    b.miniButton({ ico: 'alignTop', title: t('cmd.alignTop'), needs: ['edit', 'text'], active: (s) => s.caps.valign === 'top', onClick: () => api.cmd('setValign', 'top') }),
    b.miniButton({ ico: 'alignMiddle', title: t('cmd.alignMiddle'), needs: ['edit', 'text'], active: (s) => s.caps.valign === 'middle', onClick: () => api.cmd('setValign', 'middle') }),
    b.miniButton({ ico: 'alignBottom', title: t('cmd.alignBottom'), needs: ['edit', 'text'], active: (s) => s.caps.valign === 'bottom', onClick: () => api.cmd('setValign', 'bottom') })
  ];

  const arrangeRow = (): Widget[] => [
    b.miniButton({ ico: 'zFront', title: t('cmd.bringFront'), needs: ['edit', 'selection'], onClick: () => api.cmd('zOrder', 'front') }),
    b.miniButton({ ico: 'zFront', title: t('cmd.bringForward'), needs: ['edit', 'selection'], onClick: () => api.cmd('zOrder', 'forward') }),
    b.miniButton({ ico: 'zBack', title: t('cmd.sendBackward'), needs: ['edit', 'selection'], onClick: () => api.cmd('zOrder', 'backward') }),
    b.miniButton({ ico: 'zBack', title: t('cmd.sendBack'), needs: ['edit', 'selection'], onClick: () => api.cmd('zOrder', 'back') }),
    b.miniButton({ ico: 'distH', title: t('cmd.distH'), needs: ['edit', 'multi'], onClick: () => api.cmd('distribute', 'h') }),
    b.miniButton({ ico: 'distV', title: t('cmd.distV'), needs: ['edit', 'multi'], onClick: () => api.cmd('distribute', 'v') })
  ];

  const tabs: RibbonTab[] = [];
  const add = (id: string, label: string, groups: HTMLElement[]) => {
    const editOnly = groups.every((g) => g.classList.contains('edit-only'));
    tabs.push({ id, label, panel: h('div', { class: 'rb-panel' }, groups), editOnly });
  };

  // ---------- 开始 / Home ----------
  add('home', t('tab.home'), [
    b.group(
      t('grp.clipboard'),
      [
        [
          b.splitButton({
            ico: 'paste',
            label: t('cmd.paste'),
            accel: 'mod+v',
            needs: 'edit',
            onClick: () => api.cmd('paste'),
            menu: () => [{ label: t('cmd.paste'), icon: 'paste', accel: kbd('mod+v'), onClick: () => api.cmd('paste') }]
          })
        ],
        [
          b.miniButton({ ico: 'cut', title: withAccel(t('cmd.cut'), 'mod+x'), needs: ['edit', 'selection'], onClick: () => api.cmd('cut') }),
          b.miniButton({ ico: 'copy', title: withAccel(t('cmd.copy'), 'mod+c'), needs: ['edit', 'selection'], onClick: () => api.cmd('copy') }),
          b.miniButton({ ico: 'duplicate', title: withAccel(t('cmd.duplicate'), 'mod+d'), needs: ['edit', 'selection'], onClick: () => api.cmd('duplicate') }),
          b.miniButton({ ico: 'deleteSlide', title: withAccel(t('cmd.delete'), 'delete'), needs: ['edit', 'selection'], onClick: () => api.cmd('delete') })
        ]
      ],
      true
    ),
    b.group(
      t('grp.slides'),
      [
        [b.splitButton({ ico: 'newSlide', label: t('cmd.newSlide'), needs: 'edit', onClick: () => api.cmd('newSlide'), menu: () => layoutMenu((id) => api.cmd('newSlide', id)) })],
        [
          b.miniButton({ ico: 'layout', title: t('cmd.layout'), needs: 'edit', onClick: () => api.cmd('openLayoutMenu') }),
          b.miniButton({ ico: 'duplicateSlide', title: t('cmd.duplicateSlide'), needs: 'edit', onClick: () => api.cmd('duplicateSlide') }),
          b.miniButton({ ico: 'deleteSlide', title: t('cmd.deleteSlide'), needs: ['edit', 'manySlides'], onClick: () => api.cmd('deleteSlide') }),
          b.miniButton({ ico: 'hideSlide', title: t('cmd.hideSlide'), needs: 'edit', active: (s) => s.slideHidden, onClick: () => api.cmd('toggleSlideHidden') })
        ],
        [
          b.miniButton({ ico: 'up', title: t('cmd.slideUp'), needs: ['edit', 'notFirstSlide'], onClick: () => api.cmd('slideUp') }),
          b.miniButton({ ico: 'down', title: t('cmd.slideDown'), needs: ['edit', 'notLastSlide'], onClick: () => api.cmd('slideDown') })
        ]
      ],
      true
    ),
    b.group(
      t('grp.font'),
      [
        [
          b.selectBox({ label: t('insp.font'), options: fontOptions, width: 120, needs: ['edit', 'text'], onChange: (v) => api.cmd('setFontFace', v), onSync: (s) => String(s.caps.fontFace || '') }),
          b.selectBox({ label: t('insp.fontSize'), options: sizeOptions, width: 58, needs: ['edit', 'text'], onChange: (v) => api.cmd('setFontSize', v), onSync: (s) => String(s.caps.fontSize || '') }),
          b.colorBox({ title: t('cmd.fontColor'), ico: 'fontColor', needs: ['edit', 'text'], onInput: (v) => api.cmd('setTextColor', v), onSync: (s) => String(s.caps.textColor || '') })
        ],
        [
          b.miniButton({ glyph: 'B', glyphStyle: 'font-weight:700', title: withAccel(t('cmd.bold'), 'mod+b'), needs: ['edit', 'text'], active: (s) => !!s.caps.bold, onClick: () => api.cmd('toggleBold') }),
          b.miniButton({ glyph: 'I', glyphStyle: 'font-style:italic;font-family:Georgia,serif', title: withAccel(t('cmd.italic'), 'mod+i'), needs: ['edit', 'text'], active: (s) => !!s.caps.italic, onClick: () => api.cmd('toggleItalic') }),
          b.miniButton({ glyph: 'U', glyphStyle: 'text-decoration:underline', title: withAccel(t('cmd.underline'), 'mod+u'), needs: ['edit', 'text'], active: (s) => !!s.caps.underline, onClick: () => api.cmd('toggleUnderline') }),
          b.miniButton({ ico: 'bullets', title: t('cmd.bullets'), needs: ['edit', 'text'], active: (s) => !!s.caps.bullet, onClick: () => api.cmd('toggleBullet') }),
          b.selectBox({ label: t('cmd.lineSpacing'), options: LINE_SPACING.map((v) => ({ value: v, label: v })), width: 54, needs: ['edit', 'text'], onChange: (v) => api.cmd('setLineSpacing', v), onSync: (s) => String(s.caps.lineSpacing || '') })
        ]
      ],
      true
    ),
    b.group(t('grp.paragraph'), [textAlignRow()], true),
    b.group(
      t('grp.drawing'),
      [
        [
          b.splitButton({ ico: 'shape', label: t('cmd.shapes'), needs: 'edit', onClick: () => api.cmd('shapes'), menu: () => api.cmd('shapesMenu') }),
          b.colorBox({ title: t('cmd.fillColor'), ico: 'fillColor', needs: ['edit', 'shape'], onInput: (v) => api.cmd('setFill', v), onSync: (s) => String(s.caps.fillColor || '') }),
          b.colorBox({ title: t('cmd.lineColor'), ico: 'lineColor', needs: ['edit', 'shape'], onInput: (v) => api.cmd('setLineColor', v), onSync: (s) => String(s.caps.lineColor || '') }),
          b.miniButton({ ico: 'flipH', title: t('cmd.flipH'), needs: ['edit', 'shape'], onClick: () => api.cmd('flip', 'h') }),
          b.miniButton({ ico: 'flipV', title: t('cmd.flipV'), needs: ['edit', 'shape'], onClick: () => api.cmd('flip', 'v') })
        ],
        arrangeRow(),
        [
          b.miniButton({ ico: 'group', title: withAccel(t('cmd.group'), 'mod+g'), needs: ['edit', 'multi'], onClick: () => api.cmd('groupSel') }),
          b.miniButton({ ico: 'ungroup', title: withAccel(t('cmd.ungroup'), 'mod+shift+g'), needs: ['edit', 'ungroupable'], onClick: () => api.cmd('ungroupSel') }),
          b.miniButton({ ico: 'lock', title: t('cmd.lock'), needs: ['edit', 'selection'], active: (s) => !!s.caps.locked, onClick: () => api.cmd('toggleLock') }),
          b.miniButton({ ico: 'hideSlide', title: t('cmd.hide'), needs: ['edit', 'selection'], active: (s) => !!s.caps.elementHidden, onClick: () => api.cmd('toggleHidden') })
        ]
      ],
      true
    )
  ]);

  // ---------- 插入 / Insert ----------
  add('insert', t('tab.insert'), [
    b.group(t('grp.content'), [
      [
        b.bigButton({ ico: 'text', label: t('cmd.textBox'), needs: 'edit', onClick: () => api.cmd('textBox') }),
        b.bigButton({ ico: 'shape', label: t('cmd.shapes'), needs: 'edit', onClick: () => api.cmd('shapes') }),
        b.bigButton({ ico: 'picture', label: t('cmd.picture'), needs: 'edit', onClick: () => api.cmd('picture') })
      ]
    ], true),
    b.group(t('grp.media'), [
      [
        b.bigButton({ ico: 'table', label: t('cmd.table'), needs: 'edit', onClick: () => api.cmd('table') }),
        b.bigButton({ ico: 'chart', label: t('cmd.chart'), needs: 'edit', onClick: () => api.cmd('chart') }),
        b.bigButton({ ico: 'media', label: t('cmd.media'), needs: 'edit', onClick: () => api.cmd('media') })
      ]
    ], true),
    b.group(t('grp.slides'), [
      [b.bigButton({ ico: 'newSlide', label: t('cmd.newSlide'), needs: 'edit', onClick: () => api.cmd('newSlide') })]
    ], true)
  ]);

  // ---------- 设计 / Design ----------
  add('design', t('tab.design'), [
    b.group(
      t('grp.themes'),
      [
        [
          b.gallery(
            THEMES as any[],
            (th: any) => {
              const el = h('button', { class: 'rb-theme', type: 'button', title: data('theme', th.id, th.name) }, [
                h('span', { class: 'rb-theme-sw' }, [
                  h('i', { style: `background:${th.bg};border-color:${th.accent}` }),
                  h('i', { style: `background:${th.accent}` }),
                  h('i', { style: `background:${(th.accents || [])[1] || th.accent}` }),
                  h('i', { style: `background:${(th.accents || [])[2] || th.accent}` })
                ]),
                h('span', { class: 'rb-theme-name' }, [data('theme', th.id, th.name)])
              ]);
              el.onclick = () => api.cmd('setTheme', th.id);
              return { el, isActive: (s: RibbonState) => s.theme === th.id };
            },
            { needs: 'edit', caption: t('cmd.theme') }
          )
        ]
      ],
      true
    ),
    b.group(
      t('grp.slideSize'),
      [
        [b.selectBox({ label: t('insp.size'), options: sizeKeyOptions, width: 132, needs: 'edit', onChange: (v) => api.cmd('setSizeKey', v), onSync: (s) => s.sizeKey })],
        [b.miniButton({ ico: 'slideSize', title: t('insp.slideProps'), needs: 'edit', onClick: () => api.cmd('toggleInspector') })]
      ],
      true
    ),
    b.group(
      t('grp.background'),
      [
        [b.colorBox({ title: t('cmd.background'), ico: 'background', needs: 'edit', onInput: (v) => api.cmd('setSlideBg', v), onSync: (s) => String(s.caps.slideBg || '') })],
        [
          b.miniButton({ ico: 'picture', title: t('insp.backgroundImage'), needs: 'edit', onClick: () => api.cmd('bgImage') }),
          b.miniButton({ ico: 'background', title: t('insp.backgroundGradient'), needs: 'edit', onClick: () => api.cmd('bgGradient') }),
          b.miniButton({ ico: 'close', title: t('insp.backgroundNone'), needs: 'edit', onClick: () => api.cmd('bgNone') })
        ]
      ],
      true
    )
  ]);

  // ---------- 切换 / Transitions ----------
  add('transitions', t('tab.transitions'), [
    b.group(
      t('grp.transition'),
      [
        [
          b.gallery(
            TRANSITIONS as any[],
            (tr: any) =>
              b.chip(
                data('transition', tr.value, tr.name),
                data('transition', tr.value, tr.name),
                () => api.cmd('setTransition', tr.value),
                (s) => s.transition.type === tr.value
              ),
            { needs: 'edit', caption: t('cmd.transition') }
          )
        ]
      ],
      true
    ),
    b.group(
      t('grp.transitionOptions'),
      [
        [b.selectBox({ label: t('insp.transSpeed'), options: speedOptions, width: 78, needs: 'edit', onChange: (v) => api.cmd('setTransSpeed', v), onSync: (s) => String(s.transition.duration || DEFAULT_TRANS_SPEED) })],
        [b.advanceCheck()],
        [b.miniButton({ ico: 'add', title: t('cmd.applyTransitionAll'), needs: ['edit', 'manySlides'], onClick: () => api.cmd('applyTransitionAll') })]
      ],
      true
    )
  ]);

  // ---------- 动画 / Animations ----------
  const animTypeBox = b.selectBox({
    label: t('insp.animType'),
    options: typeOptionsFor(anim.cls),
    width: 114,
    needs: ['edit', 'single'],
    onSync: (s) => s.animPick.type,
    // 同时写回当前已选中的动画：只改草稿的话，播放效果永远停留在最初添加的那一条
    onChange: (v) => { anim.type = v; api.cmd('animPatch', { type: v }); }
  });
  let lastAnimCls = '';
  add('animations', t('tab.animations'), [
    b.group(
      t('grp.animation'),
      [
        [
          b.gallery(
            ANIM_PRESETS,
            (p) => b.chip(t(p.label), t(p.label), () => addAnim({ cls: p.cls, type: p.type })),
            { needs: ['edit', 'single'], caption: t('insp.animAdd') }
          )
        ]
      ],
      true
    ),
    b.group(
      t('grp.animationOptions'),
      [
        [
          b.selectBox({
            label: t('insp.animClass'),
            options: classOptions,
            width: 96,
            needs: ['edit', 'single'],
            onSync: (s) => {
              if (s.animPick.cls !== lastAnimCls) {
                lastAnimCls = s.animPick.cls;
                animTypeBox.setOptions(typeOptionsFor(s.animPick.cls));
              }
              return s.animPick.cls;
            },
            onChange: (v) => {
              anim.cls = v;
              const opts = typeOptionsFor(v);
              anim.type = opts[0]?.value || anim.type;
              animTypeBox.setOptions(opts);
              (animTypeBox.el.querySelector('select') as HTMLSelectElement).value = anim.type;
              api.cmd('animPatch', { presetClass: v, type: anim.type });
            }
          }),
          animTypeBox,
          b.selectBox({
            label: t('insp.animTrigger'), options: triggerOptions, width: 108, needs: ['edit', 'single'],
            onSync: (s) => s.animPick.trigger,
            onChange: (v) => { anim.trigger = v; api.cmd('animPatch', { trigger: v }); }
          }),
          b.selectBox({
            label: t('insp.animDirection'), options: directionOptions, width: 104, needs: ['edit', 'single'],
            onSync: (s) => s.animPick.direction,
            onChange: (v) => { anim.direction = v; api.cmd('animPatch', { direction: v }); }
          }),
          b.selectBox({
            label: t('insp.animDuration'),
            options: ANIM_DURATIONS.map((n) => ({ value: String(n), label: String(n) })),
            width: 66,
            needs: ['edit', 'single'],
            onSync: (s) => s.animPick.duration,
            onChange: (v) => { anim.duration = Number(v); api.cmd('animPatch', { duration: Number(v) }); }
          })
        ]
      ],
      true
    ),
    b.group(t('grp.animationList'), [[b.animListBox()]], true)
  ]);

  // ---------- 幻灯片放映 / Slide Show ----------
  add('slideshow', t('tab.slideshow'), [
    b.group(t('grp.startShow'), [
      [
        b.bigButton({ ico: 'present', label: t('cmd.present'), accel: 'f5', onClick: () => api.cmd('present') }),
        b.bigButton({ ico: 'presentHere', label: t('cmd.presentHere'), accel: 'shift+f5', onClick: () => api.cmd('presentHere') }),
        b.bigButton({ ico: 'endShow', label: t('cmd.endShow'), needs: 'presenting', onClick: () => api.cmd('endShow') })
      ]
    ]),
    b.group(t('grp.showSettings'), [
      [b.miniButton({ ico: 'grid', title: t('cmd.grid'), active: (s) => s.grid, onClick: () => api.cmd('grid') })],
      [b.miniButton({ ico: 'snap', title: t('cmd.snap'), active: (s) => s.snap, onClick: () => api.cmd('snap') })],
      [b.miniButton({ ico: 'notes', title: t('insp.notes'), needs: 'edit', onClick: () => api.cmd('toggleInspector') })]
    ])
  ]);

  // ---------- 视图 / View ----------
  add('view', t('tab.view'), [
    b.group(t('grp.window'), [
      [b.miniButton({ ico: 'edit', title: t('cmd.modeEdit'), needs: 'preview', onClick: () => api.cmd('setMode', 'edit') })],
      [b.miniButton({ ico: 'showSlide', title: t('cmd.modePreview'), needs: 'edit', onClick: () => api.cmd('setMode', 'preview') })],
      [b.miniButton({ ico: 'props', title: t('cmd.docProps'), onClick: () => api.cmd('docProps') })]
    ]),
    b.group(t('grp.window'), [
      [b.miniButton({ ico: 'grid', title: t('cmd.grid'), needs: 'edit', active: (s) => s.grid, onClick: () => api.cmd('grid') })],
      [b.miniButton({ ico: 'snap', title: t('cmd.snap'), needs: 'edit', active: (s) => s.snap, onClick: () => api.cmd('snap') })],
      [b.miniButton({ ico: 'notes', title: t('grp.window'), needs: 'edit', onClick: () => api.cmd('toggleInspector') })]
    ]),
    b.group(t('grp.export'), [
      [
        b.bigButton({ ico: 'picture', label: t('cmd.exportImage'), onClick: () => api.cmd('exportImage') }),
        b.bigButton({ ico: 'download', label: t('cmd.exportImages'), onClick: () => api.cmd('exportImages') }),
        b.bigButton({ ico: 'save', label: t('cmd.exportPdf'), onClick: () => api.cmd('exportPdf') }),
        b.bigButton({ ico: 'swap', label: t('cmd.convertFormat'), onClick: () => api.cmd('convertFormat') })
      ]
    ])
  ]);

  // ---------- 帮助 / Help ----------
  add('help', t('tab.help'), [
    b.group(t('grp.help'), [
      [
        b.bigButton({ ico: 'help', label: t('cmd.shortcuts'), onClick: () => api.cmd('shortcuts') }),
        b.bigButton({ ico: 'props', label: t('cmd.docProps'), onClick: () => api.cmd('docProps') })
      ]
    ])
  ]);

  return tabs;
}