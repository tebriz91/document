export interface WebviewMessage<T = unknown> {
  type: string;
  payload?: T;
}

export interface FileChunkPayload {
  type: string;
  chunkIndex: number;
  totalChunks: number;
  name: string;
  size: number;
  lastModified: number;
  data: string;
}
