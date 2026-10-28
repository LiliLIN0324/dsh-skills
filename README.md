# DSH / Agent Skills 汇总库

林丽丽本机（DeepSeek Harness）可用的 **95 个 Agent Skill**，按用途分类整理。每个 skill 是一个目录，核心是带 YAML frontmatter 的 `SKILL.md`（name + description + 使用说明），部分还带 `references/`、`scripts/`。

## 怎么用

- 在支持 Agent Skills 的工具里（DSH / Claude Code / Codex 等）把目录放进 skills 路径即可；`npx skills add <repo>` 也能整包安装。
- 模型按任务自动匹配 skill，也可以在任务里点名：例如「用 godot-gdscript 帮我写这段脚本」。
- `skills.json` 是机器可读清单（name / category / dir / description）。
- 每个分类目录下的 `README.md` 是该类的条目索引。

## 分类概览

| 类别 | 数量 | 目录 |
|---|---|---|
| 前端设计 | 2 | `frontend/` |
| 其他引擎（three.js / Phaser / PixiJS / Bevy / LÖVE / pygame） | 9 | `game/engines/` |
| 玩法类型模板 | 9 | `game/genres/` |
| Godot 4.7 | 16 | `game/godot/` |
| Roblox | 7 | `game/roblox/` |
| 通用游戏系统与发布 | 20 | `game/systems/` |
| Unity 6.3 | 8 | `game/unity/` |
| Unreal Engine 5 | 6 | `game/unreal/` |
| Web 动效 | 5 | `motion/` |
| ACE Studio 音乐制作 | 4 | `music/ace-studio/` |
| 办公文档（Word / Excel / PPT / PDF） | 7 | `office/` |
| 报告出图管线 | 2 | `reporting/` |
| **合计** | **95** | |

## 全部条目

| Skill | 类别 | 说明 |
|---|---|---|
| `frontend-design` | 前端设计 | Guidance for distinctive, intentional visual design when building new UI or reshaping an existing one. |
| `vercel-react-view-transitions` | 前端设计 | Guide for implementing smooth, native-feeling animations using React's View Transition API (`<ViewTransition>` component, `addTransitionType`, and CSS… |
| `bevy-ecs` | 其他引擎（three.js / Phaser / PixiJS / Bevy / LÖVE / pygame） | Structure a Bevy app around its Entity Component System: build the App with plugins, define Component/Resource types, write systems with Query/Res/Com… |
| `love2d-core` | 其他引擎（three.js / Phaser / PixiJS / Bevy / LÖVE / pygame） | Structure and debug a LÖVE (Love2D) game in Lua: the love.load/update/draw loop, delta-time movement, input, and screen states. |
| `phaser-arcade-physics` | 其他引擎（three.js / Phaser / PixiJS / Bevy / LÖVE / pygame） | Use Phaser 4 Arcade Physics: enable the world, give sprites bodies, set velocity/acceleration/gravity, and resolve collisions with colliders, overlaps… |
| `phaser-core` | 其他引擎（three.js / Phaser / PixiJS / Bevy / LÖVE / pygame） | Set up and debug a Phaser 4 game: the Game config, the Scene lifecycle (init/preload/create/update), the asset loader, cameras, and cross-scene commun… |
| `pixijs-rendering` | 其他引擎（three.js / Phaser / PixiJS / Bevy / LÖVE / pygame） | Build a PixiJS v8 render layer: create the async Application, load textures with Assets, compose the scene graph with Container and Sprite, drive the … |
| `pygame-core` | 其他引擎（three.js / Phaser / PixiJS / Bevy / LÖVE / pygame） | Structure a pygame (pygame-ce) game in Python: the init/event/update/draw loop, delta-time movement, Surface/Rect blitting, keyboard/mouse input, and … |
| `threejs-gltf-loading` | 其他引擎（three.js / Phaser / PixiJS / Bevy / LÖVE / pygame） | Load glTF/GLB models in three.js with GLTFLoader and play their skinned animations with AnimationMixer, including DRACO/Meshopt-compressed meshes and … |
| `threejs-materials-lighting` | 其他引擎（three.js / Phaser / PixiJS / Bevy / LÖVE / pygame） | Light and shade a three.js scene: choose materials (MeshStandardMaterial PBR vs unlit MeshBasicMaterial), add ambient/hemisphere/directional/point/spo… |
| `threejs-scene-setup` | 其他引擎（three.js / Phaser / PixiJS / Bevy / LÖVE / pygame） | Stand up a three.js scene: import maps and the three/addons path, the Scene/PerspectiveCamera/WebGLRenderer trio, the setAnimationLoop render loop, re… |
| `card-game` | 玩法类型模板 | Build a card game: card data, deck/hand/discard zones, draw/shuffle/reshuffle, a turn structure, costs, and effect resolution. |
| `fps-shooter` | 玩法类型模板 | Build a first-person shooter: move+mouse-look controller, hitscan or projectile shooting, weapons, health, and enemy AI. |
| `platformer` | 玩法类型模板 | Build a 2D platformer: run/jump control with coyote time, jump buffering, and variable jump height, plus tiled levels and hazards. |
| `puzzle` | 玩法类型模板 | Build a puzzle game: grid/board state, move input, rule-based resolution (match-3 cascades, sokoban pushes, tile logic), scoring, and undo. |
| `roguelike` | 玩法类型模板 | Build a roguelike: turn-based grid movement, procedural dungeons, permadeath, field-of-view, and loot tables. |
| `rpg` | 玩法类型模板 | Build an RPG: stats and leveling, inventory and equipment, quests, branching dialogue, save/load, and combat. |
| `survival-crafting` | 玩法类型模板 | Build a survival-crafting game: resource gathering, inventory, crafting and a tech tree, needs (hunger/thirst/temperature), and base building. |
| `tower-defense` | 玩法类型模板 | Build a tower defense: enemies pathing along lanes, wave spawning, towers that auto-target and fire, an economy, and lives. |
| `visual-novel` | 玩法类型模板 | Build a visual novel: a branching script, character and background display, a text box with choices, save/load, backlog, and skip/auto. |
| `godot-2d-movement` | Godot 4.7 | Implement 2D kinematic character movement in Godot 4.7 with CharacterBody2D and move_and_slide(): platformer run/jump with gravity, top-down 8-directi… |
| `godot-3d-essentials` | Godot 4.7 | Set up a Godot 4.7 3D scene: Node3D transforms, Camera3D, lighting (DirectionalLight3D/OmniLight3D), WorldEnvironment for sky/ambient/tonemap/post, Me… |
| `godot-animation` | Godot 4.7 | Animate in Godot 4.7 three ways: AnimationPlayer for keyframed clips (incl. |
| `godot-audio` | Godot 4.7 | Play and mix audio in Godot 4.7: AudioStreamPlayer (2D/3D variants), audio buses with volume/mute and effects, music vs SFX routing, db/linear volume,… |
| `godot-csharp` | Godot 4.7 | Use C#/.NET in Godot 4.7: partial classes extending nodes, the PascalCase lifecycle (_Ready/_Process/_PhysicsProcess), [Export] fields, [Signal] deleg… |
| `godot-export` | Godot 4.7 | Export and build a Godot 4.7 project for distribution: install export templates, define export presets (Windows/macOS/Linux/Web/Android), run headless… |
| `godot-gdscript` | Godot 4.7 | Write idiomatic GDScript for Godot 4.7: static typing, the node lifecycle (_ready/_process/_physics_process), @export/@onready/@tool annotations, sign… |
| `godot-gdscript-headless-testing` | Godot 4.7 | Run GDScript test suites from the command line with `godot --headless`, using a SceneTree/MainLoop runner script that exits non-zero on failure so CI … |
| `godot-multiplayer` | Godot 4.7 | Build networked games with Godot 4.7 high-level multiplayer: set up an ENetMultiplayerPeer server/client, define RPCs with the @rpc annotation (call v… |
| `godot-nodes-scenes` | Godot 4.7 | Structure a Godot 4.7 project with the scene tree and node composition: build reusable scenes, instance PackedScenes at runtime, navigate the tree saf… |
| `godot-physics` | Godot 4.7 | Use Godot 4.7 physics bodies and detection in 2D and 3D: RigidBody, StaticBody, Area, and CharacterBody; collision layers vs masks; contact/overlap si… |
| `godot-resources` | Godot 4.7 | Design data-driven Godot 4.7 games with custom Resource classes: define typed data with class_name + @export, save/load .tres/.res files, instance and… |
| `godot-shaders` | Godot 4.7 | Write Godot 4.7 shaders in the Godot Shading Language: canvas_item shaders for 2D and spatial shaders for 3D, with vertex/fragment functions, uniforms… |
| `godot-signals-groups` | Godot 4.7 | Build event-driven, decoupled Godot 4.7 gameplay with signals and node groups: declare and emit custom signals, connect with Callables (incl. |
| `godot-tilemap` | Godot 4.7 | Build and edit tile-based 2D levels in Godot 4.7 with TileMapLayer and TileSet: paint layers, set up collision/navigation/custom-data on tiles, autoti… |
| `godot-ui-control` | Godot 4.7 | Build Godot 4.7 user interfaces with Control nodes: anchors and offsets for responsive layout, Container nodes (VBox/HBox/Grid/Margin) for automatic a… |
| `roblox-characters` | Roblox | Build respawn-safe Roblox character systems around Players, CharacterAdded/CharacterRemoving, Humanoid, HumanoidRootPart, Animator, R6/R15 rigs, movem… |
| `roblox-datastores` | Roblox | Persist player data in Roblox with DataStoreService: GetDataStore, GetAsync/ SetAsync/UpdateAsync/IncrementAsync wrapped in pcall, load-on-join and sa… |
| `roblox-luau` | Roblox | Script a Roblox experience in Luau: get services, create and parent Instances, connect events, run server Scripts vs client LocalScripts, and communic… |
| `roblox-networking` | Roblox | Design and harden Roblox client/server networking with RemoteEvent, RemoteFunction, and UnreliableRemoteEvent; server authority, argument and Instance… |
| `roblox-physics` | Roblox | Implement Roblox physical simulation and queries with assemblies, anchoring, constraints, collision groups, CanCollide/CanTouch/CanQuery, raycasts and… |
| `roblox-studio-workflow` | Roblox | Modify living Roblox Studio projects safely: inspect Explorer and execution locations, preserve existing structure, reuse modules/remotes, create mean… |
| `roblox-ui` | Roblox | Build production Roblox interfaces with ScreenGui/PlayerGui lifecycle, responsive UDim2 layouts, safe insets, reusable editable Instances, cross-devic… |
| `ai-behavior-trees-utility-ai` | 通用游戏系统与发布 | Build a production behavior-tree runtime (Blackboard, action/condition leaves, sequence/selector/parallel composites, decorators) and a Utility AI sys… |
| `audio-design` | 通用游戏系统与发布 | Implement game audio practice — bus/mixer architecture and gain in decibels, ducking (sidechain), adaptive/dynamic music via layering and re-sequencin… |
| `camera-systems` | 通用游戏系统与发布 | Build game cameras that feel good — 2D follow with a deadzone, look-ahead, smoothing, and level-bounds clamping; 3D third-person orbit with collision … |
| `create-game-assets` | 通用游戏系统与发布 | Plan, generate, source, normalize, and validate cohesive visual game assets. |
| `dialogue-systems` | 通用游戏系统与发布 | Build branching dialogue and narrative — a node/choice graph with conditions, variables, and localization hooks — and choose between authoring tools I… |
| `game-ai` | 通用游戏系统与发布 | Design NPC and enemy decision-making with finite state machines, behavior trees, steering behaviors, and A* pathfinding — engine-neutral algorithms th… |
| `game-feel` | 通用游戏系统与发布 | Add "juice" and game feel that makes actions satisfying — screen shake, hit-stop/freeze frames, tweened/eased motion, squash & stretch, knockback, and… |
| `game-jam` | 通用游戏系统与发布 | Plan and ship a game under a jam deadline: lock scope to the clock, schedule the hours, cut features, and submit on time. |
| `game-ui-ux` | 通用游戏系统与发布 | Design and build game UI/UX — HUDs, menus, and overlays — that survive every screen: anchor- based responsive layout, resolution/aspect scaling and sa… |
| `input-systems` | 通用游戏系统与发布 | Architect game input — action mapping (abstracting keys into named actions), rebinding with conflict detection and persistence, multi-device support (… |
| `itch-publish` | 通用游戏系统与发布 | Publish and update a game on itch.io: create the project page and upload builds with the butler CLI (butler push) to named channels. |
| `level-design` | 通用游戏系统与发布 | Design and build playable levels — the blockout/whitebox-to-playable workflow, player metrics and grid layout, pacing and flow (tension/rest curve), g… |
| `performance-optimization` | 通用游戏系统与发布 | Find and fix game performance problems methodically — measure with the engine profiler first, reason about the frame-time budget, locate the CPU-vs-GP… |
| `physics-tuning` | 通用游戏系统与发布 | Tune game physics for stable, good-feeling motion — fixed vs variable timestep, render interpolation, mass/gravity/drag, continuous collision detectio… |
| `procedural-gen` | 通用游戏系统与发布 | Generate game content procedurally — seeded deterministic RNG, value/Perlin/ Simplex noise for terrain and heightmaps, grid dungeon generation (rooms … |
| `prototype-fast` | 通用游戏系统与发布 | Build a playable prototype in about an hour to answer one question — is it fun? — with greybox primitives, a hard timebox, and explicit keep/kill crit… |
| `router` | 通用游戏系统与发布 | Routes any game-development request to the right specialized skill(s): it detects the engine (Godot, Unity, Unreal, Bevy, Phaser, PixiJS, three.js, LÖ… |
| `save-systems` | 通用游戏系统与发布 | Design save/load for game state — choosing what to serialize, file formats, save slots, atomic crash-safe writes, schema versioning and migration, and… |
| `shader-programming` | 通用游戏系统与发布 | Write game shaders from cross-engine fundamentals — the vertex→fragment pipeline, coordinate spaces, UV math, and common 2D/3D effects (tint, UV scrol… |
| `steam-publish` | 通用游戏系统与发布 | Publish or update a game on Steam with Steamworks and SteamPipe: configure depots and packages, upload builds with steamcmd, set a build live on a bra… |
| `unity-animation` | Unity 6.3 | Drive Unity 6.3 LTS character animation with Animator Controllers: states, transitions, parameters, blend trees, animation layers, and humanoid Avatar… |
| `unity-build-pipeline` | Unity 6.3 | Build and ship Unity 6.3 LTS players: build settings and scenes, player/quality settings, the IL2CPP vs Mono scripting backend, managed code stripping… |
| `unity-csharp-scripting` | Unity 6.3 | Write Unity 6.3 LTS C# gameplay scripts: the MonoBehaviour lifecycle (Awake/OnEnable/Start/Update/FixedUpdate/LateUpdate), GameObject and component ac… |
| `unity-input-system` | Unity 6.3 | Wire player input in Unity 6.3 LTS with the Input System package: Input Actions, action maps, the PlayerInput component, and reading values via callba… |
| `unity-navmesh` | Unity 6.3 | Add AI navigation in Unity 6.3 LTS: bake a NavMesh with the AI Navigation package (NavMeshSurface), move agents with NavMeshAgent.SetDestination, and … |
| `unity-physics` | Unity 6.3 | Set up 3D physics in Unity 6.3 LTS: Rigidbody movement and forces, colliders, triggers vs collisions, layer-based collision, raycasts, and joints. |
| `unity-scriptableobjects` | Unity 6.3 | Architect Unity 6.3 LTS data and decoupling with ScriptableObjects: config/data assets, shared runtime variables, event channels, and runtime sets/reg… |
| `unity-tilemap-2d` | Unity 6.3 | Build and script 2D tilemaps in Unity 6.3 LTS: the Grid + Tilemap components, the Tile Palette, tilemap colliders, rule tiles, and runtime SetTile/Get… |
| `unreal-behavior-trees` | Unreal Engine 5 | Build NPC AI in Unreal Engine 5 with Behavior Trees and Blackboards: composites (Selector/Sequence), tasks, decorators, services, and running the tree… |
| `unreal-blueprints` | Unreal Engine 5 | Build Unreal Engine 5 gameplay with Blueprint visual scripting: Blueprint Classes, the Event Graph and Construction Script, variables/functions/macros… |
| `unreal-cpp-gameplay` | Unreal Engine 5 | Write Unreal Engine 5 C++ gameplay code: the UCLASS/UPROPERTY/UFUNCTION reflection macros, the Gameplay Framework (GameMode, Pawn, Character, PlayerCo… |
| `unreal-enhanced-input` | Unreal Engine 5 | Set up player input in Unreal Engine 5 with Enhanced Input: Input Actions, Input Mapping Contexts, modifiers and triggers, adding the mapping context,… |
| `unreal-niagara` | Unreal Engine 5 | Create and control VFX in Unreal Engine 5 with Niagara: systems and emitters, modules and the spawn/update stages, exposed User parameters, and spawni… |
| `unreal-packaging` | Unreal Engine 5 | Package and ship an Unreal Engine 5 project: the Platforms menu Package Project flow, build configurations (Development vs Shipping), cooking content,… |
| `animate` | Web 动效 | Build an animation from scratch, making the decisions in the order that determines whether it feels right — should it animate at all, what purpose, wh… |
| `animation-vocabulary` | Web 动效 | Reverse-lookup glossary that turns a vague description of a web animation or motion effect into its exact term ("the bouncy thing when a popover opens… |
| `find-animation-opportunities` | Web 动效 | Search a codebase or UI for places that don't animate but should, and reject everything that shouldn't. |
| `improve-animations` | Web 动效 | Survey a codebase's animation and motion code as a senior motion advisor, then produce a prioritized audit and self-contained implementation plans for… |
| `review-animations` | Web 动效 | Reviews animation and motion code against a high craft bar derived from Emil Kowalski's design engineering philosophy. |
| `ace-studio-audio-plugins` | ACE Studio 音乐制作 | "How to drive the audio plugins ACE Studio hosts — third-party VST3/AU instruments and effects, and Studio's own built-in effects — through Studio's b… |
| `ace-studio-features` | ACE Studio 音乐制作 | "What ACE Studio's features are and when to reach for each one — track types and what makes sound on them, audio effects, Vocal Synth and its voice li… |
| `ace-studio-setup` | ACE Studio 音乐制作 | "Guide an agent how to connect to ACE Studio if the tools needed didn't exist (acestudio-cli or ACE Studio MCP)." |
| `ace-studio-workflows` | ACE Studio 音乐制作 | "Recommended end-to-end workflows in ACE Studio — the order to run features in, which feature to reach for at each step, and where the work has to lea… |
| `office` | 办公文档（Word / Excel / PPT / PDF） | Create, edit, inspect, or convert Office documents, spreadsheets, presentations, and PDFs. |
| `office-docx` | 办公文档（Word / Excel / PPT / PDF） | Create, read, edit, and validate Word documents (.docx). |
| `office-export-pdf` | 办公文档（Word / Excel / PPT / PDF） | Convert an Office document to a requested PDF with LibreOffice Kit. |
| `office-pptx` | 办公文档（Word / Excel / PPT / PDF） | Create, read, edit, and validate PowerPoint presentations (.pptx). |
| `office-recalculate` | 办公文档（Word / Excel / PPT / PDF） | Refresh spreadsheet formula caches and verify calculated values with LibreOffice Kit. |
| `office-render` | 办公文档（Word / Excel / PPT / PDF） | Render requested document images or inspect visual layout with LibreOffice Kit. |
| `office-xlsx` | 办公文档（Word / Excel / PPT / PDF） | Read, create, edit, and validate Excel workbooks, data, and formulas (.xlsx). |
| `deck-builder` | 报告出图管线 | 把 Word/PPT/PDF 报告稳定产出一套「同尺寸、同体积、同质感」的图片集 + PPTX。固定画布 54.98×30.92cm(16:9)，一次栅格化，python-pptx 组装，可机器校验。用于 Pivot Studio 个人项目报告（1037 拼拼这类）的批量交付。 |
| `riffle-biz-deck-jpg` | 报告出图管线 | 生成 16:9 编号 JPG 系列图（商业介绍/作品展示，10–20 张），画幅 4762×2678、`fig/<slug>/NN.jpg`。当用户要做商业介绍图、作品图集、riffle 同款图时调用。 |

## 来源与许可

- 本机来源：全局 `~/.agents/skills`（82 个）、DSH 内置 `office-skills`（7 个）、本工作区 `.dsh/skills`（4 个，ACE Studio）、团队共享 `deck-builder`（1 个）。
- 均为第三方/自建 skill 的本地副本，本机没有留存来源 manifest。**对外发布或分发前请补出处与许可**；来源不明的建议只做私有备份。
- 生成时间：2026-10-23 16:59（Asia/Shanghai）
