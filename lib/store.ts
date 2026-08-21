import JSZip from 'jszip';import type {Store} from './model';import {initialStore} from './model';
const KEY='apa_v2_store';
export function load():Store{try{return {...initialStore,...JSON.parse(localStorage.getItem(KEY)||'{}')}}catch{return initialStore}}
export function save(s:Store){localStorage.setItem(KEY,JSON.stringify(s))}
export async function fileToData(file:File){return new Promise<string>((resolve,reject)=>{const r=new FileReader();r.onload=()=>resolve(String(r.result));r.onerror=()=>reject(r.error);r.readAsDataURL(file)})}
export async function backup(s:Store){const z=new JSZip();z.file('manifest.json',JSON.stringify({app:'AI Personal Assistant',version:2,createdAt:new Date().toISOString()},null,2));z.file('workspace.json',JSON.stringify(s));return z.generateAsync({type:'blob'})}
export async function restore(file:File){const z=await JSZip.loadAsync(file);const raw=await z.file('workspace.json')?.async('string');if(!raw)throw new Error('올바른 백업 ZIP이 아닙니다.');return {...initialStore,...JSON.parse(raw)} as Store}
