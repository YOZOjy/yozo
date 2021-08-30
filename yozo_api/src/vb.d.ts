declare namespace vbide{
export const enum AddIn{

}
export const enum AddIn{

}
export const enum Addins{

}
export const enum Addins{

}
export interface CodeModule{
getCountOfLines():int
deleteLines(startLine:int,count:int):void
insertLines(startLine:int,str:string):void
addFromString(str:string):void
}
export interface CodeModule{
getCountOfLines():int
deleteLines(startLine:int,count:int):void
insertLines(startLine:int,str:string):void
addFromString(str:string):void
}
export const enum CodePane{

}
export const enum CodePane{

}
export const enum CodePanes{

}
export const enum CodePanes{

}
export const enum CommandBarEvents{

}
export const enum CommandBarEvents{

}
export const enum ContainedVBControls{

}
export const enum ContainedVBControls{

}
export const enum Events{

}
export const enum Events{

}
export const enum FileControlEvents{

}
export const enum FileControlEvents{

}
export interface IDEBaseImpl{
getBinder():any
getVBE():vbide.VBE
}
export interface IDEBaseImpl{
getBinder():any
getVBE():vbide.VBE
}
export const enum IDTExtensibility{

}
export const enum IDTExtensibility{

}
export const enum LinkedWindows{

}
export const enum LinkedWindows{

}
export const enum Member{

}
export const enum Member{

}
export const enum Members{

}
export const enum Members{

}
export const enum Properties{

}
export const enum Properties{

}
export const enum Property{

}
export const enum Property{

}
export const enum Reference{

}
export const enum Reference{

}
export const enum References{

}
export const enum References{

}
export const enum ReferencesEvents{

}
export const enum ReferencesEvents{

}
export const enum SelectedVBControls{

}
export const enum SelectedVBControls{

}
export const enum SelectedVBControlsEvents{

}
export const enum SelectedVBControlsEvents{

}
export interface VBComponent{
getName():string
setName(name:string):void
getType():int
activate():void
getMacroClass():any
getCodeModule():vbide.CodeModule
}
export interface VBComponent{
getName():string
setName(name:string):void
getType():int
activate():void
getMacroClass():any
getCodeModule():vbide.CodeModule
}
export interface VBComponents{
add(type:int):vbide.VBComponent
remove(vbComponent:vbide.VBComponent):void
item(obj:any):vbide.VBComponent
getCount():int
addCustom(progID:string):vbide.VBComponent
importModule(FileName:string):vbide.VBComponent
}
export interface VBComponents{
add(type:int):vbide.VBComponent
remove(vbComponent:vbide.VBComponent):void
item(obj:any):vbide.VBComponent
getCount():int
addCustom(progID:string):vbide.VBComponent
importModule(FileName:string):vbide.VBComponent
}
export const enum VBComponentsEvents{

}
export const enum VBComponentsEvents{

}
export const enum VBControl{

}
export const enum VBControl{

}
export const enum VBControls{

}
export const enum VBControls{

}
export const enum VBControlsEvents{

}
export const enum VBControlsEvents{

}
export const enum VBE{

}
export const enum VBE{

}
export const enum VBForm{

}
export const enum VBForm{

}
export const enum VBNewProjects{

}
export const enum VBNewProjects{

}
export interface VBProject{
getProtectType():int
isSaved():boolean
getVBComponents():vbide.VBComponents
}
export interface VBProject{
getProtectType():int
isSaved():boolean
getVBComponents():vbide.VBComponents
}
export const enum VBProjects{

}
export const enum VBProjects{

}
export const enum VBProjectsEvents{

}
export const enum VBProjectsEvents{

}
export const enum Window{

}
export const enum Window{

}
export const enum Windows{

}
export const enum Windows{

}
}
