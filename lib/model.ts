export type FileItem={id:string;name:string;type:string;size:number;addedAt:string;status:'ready'|'unsupported';data?:string};
export type TemplateItem={id:string;name:string;type:string;addedAt:string};
export type Task={id:string;title:string;step:number;fileIds:string[];templateId?:string;request:string;conversation:{role:'user'|'assistant';text:string}[];plan?:Plan;result?:string;createdAt:string;updatedAt:string;status:'draft'|'completed'};
export type Plan={understanding:string;materials:string[];template:string;steps:string[];resultFormat:string;questions:string[];canExecute:boolean;limitation?:string};
export type Store={files:FileItem[];templates:TemplateItem[];tasks:Task[];archive:Task[]};
export const initialStore:Store={files:[],templates:[],tasks:[],archive:[]};
