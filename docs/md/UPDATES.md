# 应用内更新与 Cloudflare R2 发布

PasteMD 从 R2 更新清单检查版本，在托盘显示下载百分比、已下载大小和平均速度。点击更新条目可查看发布说明、下载或重试；关闭窗口会保留后台下载，托盘和窗口均可取消。下载完成后由用户点击“安装并重启”，不会自动安装。

客户端默认更新地址：

```text
https://download.richqaq.cn/pastemd/latest.json
```

尚未配置 R2 或 R2 请求失败时，会回退到已有 GitHub API 镜像和官方 API；清单里的安装包优先使用 R2，网络失败后尝试系统代理、直连和 GitHub 备用地址。校验失败也会尝试备用源，所有来源均校验不通过时停止安装。当前重试从头下载，不支持跨进程续传。

## 1. Cloudflare 配置

1. 在 Cloudflare 账户中启用 **R2**，完成开通流程。创建普通、无 jurisdiction 限制的 bucket，例如 `pastemd-releases`，选择 **Standard** 存储类别。
2. 在 bucket 的 **Settings → Custom Domains** 绑定 `download.richqaq.cn`，等待状态变为 Active。域名需要在与 R2 bucket 相同的 Cloudflare 账户中管理。通过 bucket 的 Custom Domains 绑定，不要手工把 CNAME 指向 `r2.dev`。
3. 正式下载使用这个自定义域名。`r2.dev` 是限速开发入口，生产环境可保持关闭。不需要部署 Worker 或 Pages，也不需要给桌面客户端配置 CORS。
4. 在 **R2 → API Tokens → Manage** 创建 R2 S3 凭据，权限选择 **Object Read & Write**，仅授权 `pastemd-releases` bucket。保存 **Access Key ID**、**Secret Access Key** 和账户 **Account ID**。这里使用的是 S3 凭据，不是普通 Cloudflare API token 的字符串。
5. 下载子域名不能要求 Cloudflare Access 登录，也不能对更新接口和安装包显示浏览器人机验证挑战；桌面客户端无法完成这类挑战。
6. 为更新清单添加 **Cache Rule → Eligible for cache**，尊重源站缓存头，或将 Edge TTL 设置为 5 分钟。脚本写入 `Cache-Control: public, max-age=300`；客户端不发送强制绕过缓存的请求头。JSON 默认不缓存，必须显式配置规则。`/pastemd/releases/` 下的 `.exe`、`.dmg` 设置可缓存，脚本写入一年缓存及 `immutable`，禁止修改相同版本路径下的文件内容。Free/Pro/Business 单文件缓存上限是 512 MB，超过上限需另行处理，不能假定安装包已经命中缓存。
7. 可对下载域名配置 WAF 限流，使用 Block，不使用客户端无法完成的浏览器挑战。设置账单告警；告警不会自动停止服务。异常时可在 bucket Settings 禁用公开域名访问；`R2_ENABLED=false` 仅停止以后的 CI 上传，不能阻止既有公开地址被访问。

使用其他域名/对象前缀时，需要保持 GitHub Variables 与客户端更新地址一致，见下文。

R2 Standard 每月包含 10 GB-month 存储、100 万次 Class A 操作、1,000 万次 Class B 操作，直接从 R2 下载的出站流量不收费；超出存储/请求额度仍会计费。先保留少量历史版本，关注 R2 用量。免费 Cloudflare 全球网络不保证大陆下载质量，配置后应在大陆网络实测。

参考：[R2 开通](https://developers.cloudflare.com/r2/get-started/)、[S3 凭据](https://developers.cloudflare.com/r2/api/tokens/)、[公开下载与域名](https://developers.cloudflare.com/r2/buckets/public-buckets/)、[价格](https://developers.cloudflare.com/r2/pricing/)。

## 2. GitHub 配置

在仓库 **Settings → Secrets and variables → Actions** 中填写以下内容。

**Secrets：**

| 名称 | 值 |
| --- | --- |
| `R2_ACCOUNT_ID` | Cloudflare Account ID |
| `R2_ACCESS_KEY_ID` | R2 S3 Access Key ID |
| `R2_SECRET_ACCESS_KEY` | R2 S3 Secret Access Key |

**Variables：**

| 名称 | 示例 / 说明 |
| --- | --- |
| `R2_ENABLED` | `true`，全部配置完成后再启用；未启用时仍发布 GitHub Release |
| `R2_BUCKET` | `pastemd-releases` |
| `R2_PUBLIC_BASE_URL` | `https://download.richqaq.cn`，不包含 `/pastemd`，不要填 S3 endpoint |
| `R2_KEY_PREFIX` | `pastemd`，不填也默认使用此值 |

上传凭据只被 GitHub Actions 使用，不会被打包到客户端或写进更新清单。普通上传无需 Cloudflare Zone ID、Global API Key 或 Worker token。

现有 macOS 构建的 Secrets 仍然需要保留：

- `MACOS_CERTIFICATE_P12`：Developer ID Application 证书的 P12 文件 Base64。
- `MACOS_CERTIFICATE_PASSWORD`：P12 密码。
- `MACOS_KEYCHAIN_PASSWORD`：临时 keychain 密码，可选。
- `APPLE_ID`、`APPLE_TEAM_ID`、`APPLE_APP_SPECIFIC_PASSWORD`：Apple 公证所需凭据。
- `NOTARY_PROFILE`：可选，默认 `PasteMDNotary`。

Workflow 已声明 `contents: write`；确认仓库/组织策略允许 GitHub Actions 创建 Releases。macOS 使用 `macos-15`（arm64）和 `macos-15-intel`（x86_64）分别构建；Windows 生成 x64 Inno Setup 安装包。

## 3. 发布流程

1. 修改 `pastemd/__init__.py` 中的 `__version__`，例如改为 `0.1.7.7`。
2. 提交发布代码，再创建与源码版本一致的标签 `v0.1.7.7` 并推送。标签不一致时工作流会提前失败。
3. Actions 完成 Windows、macOS arm64 和 x86_64 构建、macOS 签名及公证。
4. 发布 GitHub Release，保存 Windows 和两个 macOS 包，并读取 Release 的发布说明。
5. `R2_ENABLED=true` 时上传 R2 包，核对对象大小和 SHA-256 元数据，然后上传该版本的清单；所有包上传成功后才写 `pastemd/latest.json`。
6. 同版本相同内容可以重试发布；同版本重建出不同内容会失败，需要递增版本后重新打标签。旧版本补发不会将 R2 的最新稳定版本回退。

对象示例：

```text
pastemd/latest.json
pastemd/releases/v0.1.7.7/manifest.json
pastemd/releases/v0.1.7.7/PasteMD_pandoc-Setup_v0.1.7.7.exe
pastemd/releases/v0.1.7.7/PasteMD-0.1.7.7-arm64.dmg
pastemd/releases/v0.1.7.7/PasteMD-0.1.7.7-x86_64.dmg
```

带 `dev`、`alpha`、`beta` 或 `rc` 的标签发布为 GitHub 预览版本，并可以上传 R2，但不会更新稳定版 `latest.json`。普通 `ci/**` 分支和其他分支的手动运行只构建包；`codex/r2-updater` 有独立的测试上传任务，见第 7 节。

**第一次发布包含更新功能的客户端时，旧版用户仍需手动安装一次。之后可以应用内更新。**

## 4. 客户端配置与安装行为

默认配置包含：

```json
{
  "update_manifest_url": "https://download.richqaq.cn/pastemd/latest.json",
  "update_channel": "stable"
}
```

更换下载域名时：修改 `pastemd/utils/update_manifest.py` 的默认地址，让新安装用户使用正确地址；已有客户端可在配置文件中修改 `update_manifest_url` 后重启。也要同步 GitHub 的 `R2_PUBLIC_BASE_URL`/`R2_KEY_PREFIX`。客户端只接受 HTTPS，安装包只能来自配置清单的同源域名或 PasteMD 的 GitHub Release 路径。

- 每个下载包必须通过完整大小与 SHA-256 校验，失败时不执行安装。
- macOS：从只读 DMG 提取对应架构应用，检查 Bundle ID、版本、当前应用的代码签名要求和系统安全评估。复制并复验成功后，辅助进程等待旧进程退出，再替换原应用；替换或启动命令失败时尝试恢复旧应用。目录不可写或正在挂载的 DMG 中运行时，提示用户先安装到可写的“应用程序”目录。恢复也失败时保留备份并记录路径，避免删除唯一的旧应用。
- Windows：仅更新与 Inno Setup 卸载注册表匹配的正式安装目录；辅助进程等待 PasteMD 退出，调用安装包更新同一目录，成功后重启。所有用户安装需要系统 UAC 授权，当前用户安装使用原权限模式。安装失败或授权取消时弹出提示、记录日志，并尝试启动安装目录中的应用。Windows 安装程序中途失败可能留下不完整安装，不能承诺自动回滚；可使用正式安装包修复。
- 文档处理尚未结束时阻止安装，提示处理结束后重试；准备退出安装时屏蔽新热键任务。
- 源码和便携版不会被应用内安装覆盖。配置和生成文档都在安装目录之外，不会主动删除。macOS 更新成功启动后清理暂存目录；安装辅助进程接管前取消会清理下载及暂存包。

## 5. 日志与排查

托盘“查看日志”打开主日志；更新窗口“打开日志目录”可查看全部更新日志。

| 系统 | 日志目录 |
| --- | --- |
| macOS | `~/Library/Logs/PasteMD/` |
| Windows | `%APPDATA%\PasteMD\` |

- `pastemd.log`：版本来源、更新状态、下载源、代理/直连尝试、校验结果、取消、异常堆栈、安装准备命令与错误。现有滚动日志策略继续生效。
- `update-install.log`：退出后辅助进程的等待、安装、替换、恢复和重启结果。
- `update-installer.log`：Windows Inno Setup 的详细安装日志。
- Actions 日志：每个上传对象的校验结果和 `latest.json` 是否更新；生成的清单保存在 `PasteMD-update-manifest` Actions artifact 中。

常见情况：

- **CF 403 或返回 HTML**：检查 Access/WAF/人机挑战是否拦截下载路径。
- **CF 404**：检查 bucket、自定义域名、`R2_KEY_PREFIX` 和对象路径是否一致；首个版本上传前 `latest.json` 尚不存在。
- **仍看到旧版**：更新清单允许缓存 5 分钟，可等待过期或在 CF 清除它的缓存；不要清除并重建安装包来“刷新”版本。
- **上传 AccessDenied**：确认 R2 token 有该 bucket 的 Object Read & Write 权限，使用的是 S3 Access Key/Secret。
- **macOS 签名失败**：后续版本需要满足当前安装应用的签名要求，保持 Bundle ID、Apple 开发者团队和发行签名一致。
- **下载取消仍短暂等待**：网络读超时为 15 秒；macOS 正在复制/验签时需等当前系统命令结束，然后清理资源。

## 6. 本地验证与首次上线验收

无需 CF 凭据的发布预演（安装包文件名与当前源码版本必须一致）：

```bash
.venv/bin/python scripts/publish_update.py \
  --tag v0.1.7.7 --artifacts artifacts --notes release-notes.txt \
  --dry-run --output /tmp/pastemd-update-manifest.json

.venv/bin/python -m pytest tests/test_updates.py -q
```

首次上线后，用实际安装版分别验证：

1. 更新清单与两个架构包/Windows 包的公开下载地址返回 200，清单 `version` 正确。
2. 托盘下载百分比、大小和速度会更新；关闭窗口仍继续；取消后可重新下载。
3. 临时切断网络时托盘显示错误或备用源进度，日志包含原因；重新联网可以重试。
4. Windows 当前用户安装与所有用户安装均可更新；取消 UAC 不会静默失败。
5. macOS 同签名的新版本更新后可启动，配置保留，热键/权限仍可使用。
6. 在大陆网络实际测量 CF 下载速度和可用性。

当前代码测试验证下载校验、备用连接、取消竞争、托盘状态、发布顺序、重复版本保护，以及临时目录内的 macOS 替换/恢复。真正的 CF 上传与两个系统的正式安装升级仍需要配置后做上述验收；替换命令回滚测试不能证明新版启动后的运行状态。

## 7. 测试分支自动打包和 R2 上传

`codex/r2-updater` 仅手动运行才打包 Windows、macOS arm64/x86_64；推送这个分支不触发构建，测试需用户明确允许后开始。CI 在临时 checkout 中将版本变为“源码数字版本最后一段加一 + dev + workflow run number”，例如源码 `0.1.7.6` 的第 123 次工作流生成 `0.1.7.7dev123`。源文件中的正式版本不提交修改；新工作流生成更高版本，方便测试 A → B。

测试包内置 `preview` 通道，更新地址是测试公开地址加 `/pastemd-test/latest-preview.json`。这个通道只接受预览清单；请求失败时不回退到正式 GitHub 版本。测试任务不创建 GitHub Release，也不更新任何正式版 `latest.json`。

### 7.1 CF 测试配置

1. 创建 Standard bucket `pastemd-updates-test`。
2. 有 CF 托管域名时，绑定独立测试子域名，例如 `download-test.example.com`，保持 `r2.dev` 关闭。主域名解析保留在阿里云且暂不迁移时，小规模测试可临时启用 **Settings → Public Development URL**，使用 CF 提供的 `https://pub-....r2.dev`。不要在阿里云把 CNAME 指向 `r2.dev`；这种接入不受支持。
3. `r2.dev` 有限速，无 CDN 缓存、WAF 和 bot 管理，测试会直接产生 R2 操作；测试后可以关闭。正式上线使用自定义域名。免费 CF 不能直接将独立子域名作为 zone 托管；独立子域名托管要求 Enterprise，保留外部权威 DNS 的 Partial Setup 要求 Business/Enterprise。域名注册商仍可保留在阿里云，迁移的是 DNS 托管，二者不同。
4. 创建 **Object Read & Write** 的 R2 S3 token，仅授权测试 bucket。保存 Account ID、Access Key ID 和 Secret Access Key，直接填入 GitHub Secrets，不放进代码或聊天。

参考：[R2 域名与开发地址](https://developers.cloudflare.com/r2/buckets/public-buckets/)、[子域名托管](https://developers.cloudflare.com/dns/zone-setups/subdomain-setup/)、[Partial Setup](https://developers.cloudflare.com/dns/zone-setups/partial-setup/)、[R2 缓存](https://developers.cloudflare.com/cache/interaction-cloudflare-products/r2/)。

### 7.2 GitHub 测试配置

在 **Settings → Secrets and variables → Actions** 设置仓库级配置。测试任务仅使用以下名字，不回退到正式 `R2_*` 凭据；token 的 bucket 权限也应仅授权测试 bucket。

| 类型 | 名称 | 值 |
| --- | --- | --- |
| Secret | `R2_TEST_ACCOUNT_ID` | CF Account ID |
| Secret | `R2_TEST_ACCESS_KEY_ID` | 测试 bucket 的 S3 Access Key ID |
| Secret | `R2_TEST_SECRET_ACCESS_KEY` | 测试 bucket 的 S3 Secret Access Key |
| Variable | `R2_TEST_BUCKET` | `pastemd-updates-test` |
| Variable | `R2_TEST_PUBLIC_BASE_URL` | bucket 的 HTTPS 公开地址，不带对象路径，如 `https://pub-....r2.dev` |
| Variable | `R2_TEST_ENABLED` | 全部配置好后填 `true` |

测试前缀固定为 `pastemd-test`，不能通过正式 `R2_KEY_PREFIX` 改写。未启用测试上传时仍构建并保存包、生成清单，但 summary 明确显示上传已跳过。实际上传后 CI 校验公开清单和三个包的 HEAD/大小；公开清单因缓存延迟时每 60 秒检查一次，最多 7 次、等待 360 秒。它不代替真实客户端下载、哈希校验和安装验收。

### 7.3 完整 A → B 测试

1. 完成上述配置并明确允许测试后，在 Actions 的 **Build release packages → Run workflow** 选择 `codex/r2-updater`。推送分支不会触发构建。不要创建正式版本标签。
2. 三个平台构建和 **Publish isolated R2 preview** 完成后，检查 summary、公开清单版本和 Actions artifacts。安装第一轮包作为 A；旧的正式客户端没有更新代码，不能作为这个测试的基线。
3. 在独立系统用户或 VM 测试，避免覆盖日常使用的应用及配置。测试包与正式版有相同 Bundle ID / Inno AppId，安装位置和配置可能共享。配置文件默认值会被已有用户配置覆盖，因此检查 `update_channel=preview` 和 `update_manifest_url=<测试公开地址>/pastemd-test/latest-preview.json`。macOS 配置在 `~/Library/Application Support/PasteMD/config.json`，Windows 在 `%APPDATA%\\PasteMD\\config.json`。
4. **启动一次新的 workflow run** 生成 B。不要使用 Re-run 覆盖同一个 dev 版本的不同字节：版本路径不可变，重建出的签名/时间戳不同会被拒绝。
5. 等 B 发布后，从 A 托盘检查更新；清单可缓存最多 5 分钟。验证进度、关窗继续、取消、断网重试、日志、校验、安装重启，以及版本已升级且配置保留。
6. 测试结束后关闭 `r2.dev` 或测试域名公开入口，设置 `R2_TEST_ENABLED=false` 停止后续自动上传。保留 token 与公开入口都独立于正式发布。
