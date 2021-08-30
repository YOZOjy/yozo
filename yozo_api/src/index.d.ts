/// <reference path="office.d.ts" />
/// <reference path="wp.d.ts" />
/// <reference path="ss.d.ts" />
/// <reference path="ppt.d.ts" />
/// <reference path="vb.d.ts" />

declare namespace yozo {
  export function WApplication(): word.Application
  export function SApplication(): excel.Application
  export function PApplication(): powerpoint.Application
}