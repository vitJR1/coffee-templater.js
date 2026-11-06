/**
 * Host side of the coffee-templater-engine WebAssembly ABI (core spec §9.1).
 *
 * Exports used: memory, ct_abi_version, ct_alloc, ct_free, ct_call. Import provided: ct_host.call.
 * One engine instance per JS realm (JS is single-threaded); it is created lazily on first use and
 * re-created after a crash. Compiled templates are bound to an instance — see `engineGeneration`.
 */
import { TemplateError, type TemplateErrorInfo } from '../errors'
import { decodeFrame, encodeFrame, type Frame } from './frame'
import { base64ToBytes } from './values'

const wasmBytes = () => inflateSync(base64ToBytes(WASM_DEFLATED_BASE64), { out: new Uint8Array(WASM_SIZE) }) as Uint8Array<ArrayBuffer>
import { inflateSync } from 'fflate'
import { ENGINE_SHA256, ENGINE_VERSION, WASM_DEFLATED_BASE64, WASM_SIZE } from './wasm-binary'

export { ENGINE_SHA256, ENGINE_VERSION }

const SUPPORTED_ABI = 1

interface Exports {
  memory: WebAssembly.Memory
  ct_abi_version(): number
  ct_alloc(len: number): number
  ct_free(ptr: number, len: number): void
  ct_call(ptr: number, len: number): number
}

/** Answers a callback from the engine: returns the response frame JSON (and blobs). */
export type HostHandler = (request: any, blobs: Uint8Array[]) => { json: string; blobs?: Uint8Array[] }

let wasmModule: WebAssembly.Module | undefined
let current: Exports | undefined
let generation = 0
let busy = false
let host: HostHandler | undefined

function compileSync(): WebAssembly.Module {
  wasmModule ??= new WebAssembly.Module(wasmBytes())
  return wasmModule
}

/**
 * Compiles the engine asynchronously. Optional: everything works without it, but in browsers
 * calling `await initEngine()` once at startup avoids compiling the WebAssembly module
 * synchronously on the main thread.
 */
export async function initEngine(): Promise<void> {
  wasmModule ??= await WebAssembly.compile(wasmBytes())
  instance()
}

/** Drops the engine instance (frees its memory). Templates recompile transparently on next use. */
export function resetEngine(): void {
  if (busy) throw new Error('coffee-templater: cannot reset the engine while it is running')
  current = undefined
}

/** Identifies the current engine instance; changes when the instance is re-created. */
export function engineGeneration(): number {
  instance()
  return generation
}

function instance(): Exports {
  if (current) return current
  const inst = new WebAssembly.Instance(compileSync(), { ct_host: { call: hostCall } })
  const ex = inst.exports as unknown as Exports
  const abi = ex.ct_abi_version()
  if (abi !== SUPPORTED_ABI) {
    throw new Error(`coffee-templater: engine ABI ${abi} is not supported (expected ${SUPPORTED_ABI})`)
  }
  current = ex
  generation++
  return ex
}

function errorFrame(message: string): Uint8Array {
  return encodeFrame(JSON.stringify({ ok: false, message }))
}

/** `ct_host.call`: the engine asks the host to run a filter / nullGetter / stringify. */
function hostCall(ptr: number, len: number): number {
  const ex = current
  if (!ex) return 0
  let frame: Uint8Array
  try {
    const req = decodeFrame(new Uint8Array(ex.memory.buffer, ptr, len))
    if (!host) frame = errorFrame('No host handler is active')
    else {
      const res = host(req.json, req.blobs)
      frame = encodeFrame(res.json, res.blobs)
    }
  } catch (e) {
    // Never let a JS exception unwind through WebAssembly frames.
    frame = errorFrame(e instanceof Error ? e.message : String(e))
  }
  const out = ex.ct_alloc(frame.length + 4)
  if (!out) return 0
  new DataView(ex.memory.buffer).setUint32(out, frame.length, true)
  new Uint8Array(ex.memory.buffer, out + 4, frame.length).set(frame)
  return out
}

/** Sends a raw request frame and returns the decoded response frame. */
export function callEngine(json: string, blobs: readonly Uint8Array[] = [], handler?: HostHandler): Frame {
  if (busy) {
    throw new Error('coffee-templater: the engine is busy; templater functions cannot be called from inside a filter, nullGetter or stringify')
  }
  const ex = instance()
  const request = encodeFrame(json, blobs)
  busy = true
  host = handler
  try {
    const reqPtr = ex.ct_alloc(request.length)
    if (!reqPtr) throw new TemplateError('The engine is out of memory', { code: 'limit_exceeded' })
    new Uint8Array(ex.memory.buffer, reqPtr, request.length).set(request)
    const resPtr = ex.ct_call(reqPtr, request.length)
    ex.ct_free(reqPtr, request.length)
    if (!resPtr) throw new TemplateError('The engine is out of memory', { code: 'limit_exceeded' })
    const resLen = new DataView(ex.memory.buffer).getUint32(resPtr, true)
    const frame = decodeFrame(new Uint8Array(ex.memory.buffer, resPtr + 4, resLen))
    ex.ct_free(resPtr, resLen + 4)
    return frame
  } catch (e) {
    if (e instanceof WebAssembly.RuntimeError) {
      // The instance is unusable after a trap (e.g. out of memory): start over next time.
      current = undefined
      throw new TemplateError(`The engine crashed: ${e.message}`, { code: 'internal' })
    }
    throw e
  } finally {
    busy = false
    host = undefined
  }
}

export interface EngineResult<T> {
  result: T
  blobs: Uint8Array[]
  warnings: TemplateErrorInfo[]
}

/** Sends a request and unwraps `{ ok, result }`, throwing `TemplateError` on failure. */
export function request<T = any>(json: string, blobs: readonly Uint8Array[] = [], handler?: HostHandler): EngineResult<T> {
  const { json: res, blobs: out } = callEngine(json, blobs, handler)
  if (!res.ok) throw TemplateError.fromList(res.errors ?? [{ code: 'internal', message: 'Unknown engine error' }])
  return { result: res.result as T, blobs: out, warnings: res.warnings ?? [] }
}

export interface EngineInfo {
  coreVersion: string
  abiVersion: number
  modelVersion: number
  formats: string[]
  features: Record<string, unknown>
}

/** Versions and capabilities of the bundled engine. */
export function engineInfo(): EngineInfo & { sha256: string } {
  return { ...request<EngineInfo>('{"op":"info"}').result, sha256: ENGINE_SHA256 }
}
