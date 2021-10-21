/// <reference path="office.d.ts" />
/// <reference path="wp.d.ts" />
/// <reference path="ss.d.ts" />
/// <reference path="ppt.d.ts" />
/// <reference path="vb.d.ts" />

declare namespace yozo {
  export function WApplication(): word.Application
  export function SApplication(): excel.Application
  export function PApplication(): powerpoint.Application

  export function showDialog(url: string, title?: string, width?: number, height?: number, isModal?: boolean): void;
  export function showTaskPane(plugin: any, taskName: string);
  export function setControlState(controlId: String, controlState: int): void;

  export interface Plugin {
    readonly descriptor: PluginDescriptor
  }
  export interface PluginDescriptor {
    readonly name: string
    readonly type: string
    readonly version: string
    readonly location: string
    readonly localLocation: string
  }
}