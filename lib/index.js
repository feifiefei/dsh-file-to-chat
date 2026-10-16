// dsh-file-to-chat 的 Node half。
//
// P1 阶段这里是空实现：通路验证只关心「浏览器端按钮 -> composer 草稿」，
// 全部逻辑住在 client.js（经 package.json 的 dsh.client 声明接入 dsh web）。
// P2 起此文件会承载真正的文件解析与 Remote 面（xlsx/docx/csv/txt）。
export const name = 'dsh-file-to-chat'

export function apply() {}

export default { name, apply }
