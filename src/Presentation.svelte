<script lang="ts">
    import { slide } from "svelte/transition";
    import PresentationVideoWord from "./assets/PresentationVideo-Word.mp4";
    import PresentationVideoExcel from "./assets/PresentationVideo-Excel.mp4";
    import { cubicInOut } from "svelte/easing";
    import { onMount } from "svelte";
    /**
     * If information about sideloading the extension should be shown
     */
    let showDownloadInfo = $state(false);
    /**
     * Download instructions for which OS
     */
    let downloadType = $state("macos");
    /**
     * Show the Excel presentation video instead of the Word one
     */
    let showExcelVideo = $state(false)
    /**
     * If there's not enough space to center the text, and therefore items must be set to scrollable
     */
    let smallSpace = $state(false);
    /**
     * Update the `smallSpace` value
     */
    function check() {
        const item = document.querySelector(".miniWidth") as HTMLElement;
        if (item?.firstElementChild) {
            smallSpace =
                item.firstElementChild.getBoundingClientRect().height >
                window.innerHeight -
                    (document.getElementById("header") as HTMLElement).getBoundingClientRect().height * 2.5
                    || (window.innerWidth / window.devicePixelRatio) < 500;
        }
    }
    onMount(() => {
        try {
            Office.onReady((host) => {
                if (host.host) window.location.href = "./?ForceOfficeUI=1";
            })
        } catch(ex) {
            console.warn(ex);
        }
        window.addEventListener("resize", () => {
            check();
        });
        setTimeout(() => check(), 100);
    });
    $effect(() => {
        if (downloadType) { // In this way, Svelte knows that this effect should run when the "downloadType variable changes"
            check();
            for (let i of [100, 150, 200, 210, 250]) setTimeout(check, i);
        }
    });
    /**
     * Download the XML file to sideload the extension
     */
    async function downloadStyle() {
        const req = await fetch(
            "https://raw.githubusercontent.com/dinoosauro/custom-office-styling/refs/heads/main/CustomOfficeStyling.xml",
        );
        const a = Object.assign(document.createElement("a"), {
            href: URL.createObjectURL(await req.blob()),
            target: "_blank",
            download: "CustomOfficeStyling.xml",
        });
        a.click();
    }
</script>

<main style="position: fixed; top: 0; left: 0; overflow: auto; transition: filter 0.35s ease-in-out">
    <div id="header" class="flex hcenter gap">
        <img
            style="width: 48px; height: 48px"
            src="./logo_light.svg"
            alt="Website icon. Click on it to go back to the Selection tab"
        />
        <h2>Custom Office Styling</h2>
    </div>
    <div
        style={smallSpace
            ? ""
            : "position: fixed; top: 0; left: 40px; width: calc(100vw - 80px)"}
    >
        {#if !showDownloadInfo}
            <div
                class={smallSpace ? "miniWidth fullWidth" : "presentation miniWidth floatLeft"}
            >
                <div>
                    <h1>
                        Customize the styling of your Word adnd Excel documents, both from
                        Desktop, Web and iPad
                    </h1>
                    <p>
                        Create some styles you can quickly apply from Word,
                        import and export them, create new shapes with custom colors and gradients, add links to specific parts of your text, or just change even the
                        slightest details of the paragraphs, tables or lists
                        you've already written. Moreover, change the properties of Excel charts, and export them as a picture. All for free.
                    </p>
                    <br />
                    <div class="flex gap">
                        <button onclick={async () => {
                            const main = document.querySelector("main");
                            if (main) {
                                main.style.filter = "blur(16px) brightness(50%)"
                                await new Promise(res => setTimeout(res, 360));
                            }
                            showDownloadInfo = true;
                            if (main) main.style.filter = "";
                        }}
                            >Download now</button
                        >
                        <button style="width: max-content; background-color: var(--input); white-space: pre" onclick={() => window.open("https://github.com/dinoosauro/custom-office-styling")}>View source code</button>
                    </div>
                </div>
            </div>
        {:else}
            <div
                class={smallSpace ? "miniWidth fullWidth" : "presentation miniWidth floatLeft"}
            >
                <div>
                    <h1>Sideload the add-in:</h1>
                    <div class="flex gap" style="overflow: auto">
                        <span
                            onclick={() => (downloadType = "macos")}
                            class="card secondCard chip"
                            style={downloadType === "macos"
                                ? "background-color: var(--accent)"
                                : undefined}
                        >
                            iPadOS / macOS
                        </span>
                        <span
                            onclick={() => (downloadType = "windows")}
                            class="card secondCard chip"
                            style={downloadType === "windows"
                                ? "background-color: var(--accent)"
                                : undefined}
                        >
                            Windows
                        </span>
                        <span
                            onclick={() => (downloadType = "web")}
                            class="card secondCard chip"
                            style={downloadType === "web"
                                ? "background-color: var(--accent)"
                                : undefined}
                        >
                            Web
                        </span>
                    </div>
                    {#if downloadType === "macos"}
                        <div
                            in:slide={{ duration: 400, easing: cubicInOut }}
                            out:slide={{ duration: 400, easing: cubicInOut }}
                        >
                            <p>
                                The sideloading process is quite easy. First,
                                download the <u onclick={downloadStyle}
                                    >CustomOfficeStyle.xml</u
                                > file. Then:
                            </p>
                            <ul>
                                <li>
                                    If you're using iPadOS, copy it in the
                                    "Word" or "Excel" folder in the Files app;
                                </li>
                                <li>
                                    If you're using macOS, copy it in the folder
                                    you can find by opening the Finder, pressing
                                    Command + Shift + G and pasting this: <code
                                        onclick={(e) =>
                                            navigator.clipboard.writeText(
                                                (e.target as HTMLElement)
                                                    .textContent,
                                            )}
                                        >./Library/Containers/com.microsoft.Word/Data/Documents/wef</code
                                    ><br>
                                    <ul>
                                        <li style="font-size: 0.9em;">If you want to sideload it to Excel, just change <code>com.microsoft.Word</code> with <code>com.microsoft.Excel</code></li>
                                    </ul>
                                </li>
                            </ul>
                            <p>Now, close Word and reopen it. You'll find the add-in in the "Add-ins" button of the Home tab.</p>
                        </div>
                    {:else if downloadType === "windows"}
                        <div
                            in:slide={{ duration: 400, easing: cubicInOut }}
                            out:slide={{ duration: 400, easing: cubicInOut }}
                        >
                            <p>
                                Sideloading on Windows is a little bit tricky.
                                Follow these steps:
                            </p>
                            <ul>
                                <li>
                                    Download the <u onclick={downloadStyle}
                                        >CustomOfficeStyle.xml</u
                                    >
                                    file and save it in the
                                    <code>CustomOfficeStyling</code> folder
                                </li>
                                <li>
                                    Right-click the <code
                                        >CustomOfficeStyling</code
                                    > folder, click on "Properties" -> "Sharing"
                                    -> "Advanced sharing"
                                </li>
                                <li>
                                    Now tick the "Share folder" checkbox, and
                                    apply the changes. This will make the
                                    content of the folder public to other users
                                    on your network, so keep only the
                                    CustomOfficeStyle file there
                                </li>
                                <li>
                                    Open the Registry editor, and go to: <code
                                        onclick={(e) =>
                                            navigator.clipboard.writeText(
                                                (e.target as HTMLElement)
                                                    .textContent,
                                            )}
                                        >HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs</code
                                    >
                                </li>
                                <li>
                                    Right-click on the folder, click on New ->
                                    Key, and navigate inside that folder
                                </li>
                                <li>
                                    Set the "Flags" proprety to "1", and set the
                                    "Url" property to the URL of the shared
                                    folder
                                </li>
                                <li>
                                    Open Word or Excel, and from the Home tab select
                                    "Add-ins". Click on "More add-ins", then in
                                    the new pop-up on "Shared folder"
                                </li>
                                <li>
                                    Select the "Custom Office Styling" extension,
                                    and install it.
                                </li>
                            </ul>
                        </div>
                    {:else if downloadType === "web"}
                        <div
                            in:slide={{ duration: 400, easing: cubicInOut }}
                            out:slide={{ duration: 400, easing: cubicInOut }}
                        >
                            <p>The sideloading process is easy:</p>
                            <ul>
                                <li>
                                    Download the <u onclick={downloadStyle}
                                        >CustomOfficeStyle.xml</u
                                    > file
                                </li>
                                <li>
                                    Click on the "Add-ins" button in the "Home
                                    tab"
                                </li>
                                <li>
                                    Click on "More add-ins", then on "My
                                    Add-ins"
                                </li>
                                <li>
                                    Click "Upload my add-in", and select the <code
                                        >CustomOfficeStyle.xml</code
                                    > file
                                </li>
                            </ul>
                        </div>
                    {/if}
                    <p>
                        This add-in is completely <a
                            href="https://github.com/dinoosauro/custom-office-styling"
                            target="_blank">open source</a
                        >.
                    </p>
                    <i style="font-size: 0.6em;">Word, Excel, Windows and Office are trademarks of Microsoft, which is in no way affiliated with this project.<br>iPad, iPadOS and macOS are trademarks of Apple, which is in no way affiliated with this project.</i>
                </div>
            </div>
        {/if}
        <div
            class={smallSpace ? "miniWidth fullWidth marginTop wcenter" : "presentation miniWidth floatRight wright"}
            style="display: flex"
        >
        <div class="wcenter" style="display: flex; flex-direction: column">
                                <div class="flex gap wcenter" style="overflow: auto; margin-bottom: 10px">
                        <span
                            onclick={() => (showExcelVideo = false)}
                            class="card secondCard chip"
                            style={!showExcelVideo
                                ? "background-color: var(--accent)"
                                : undefined}
                        >
                            Word video
                        </span>
                        <span
                            onclick={() => (showExcelVideo = true)}
                            class="card secondCard chip"
                            style={showExcelVideo
                                ? "background-color: var(--accent)"
                                : undefined}
                        >
                            Excel video
                            </span>
                            </div>
                            <div class="flex wcenter" style="width: 100%">
                                <video src={showExcelVideo ? PresentationVideoExcel : PresentationVideoWord} autoplay muted loop></video>
                            </div>
            </div>
        </div>
    </div>
</main>

<style>
    .presentation {
        display: flex;
        align-items: center;
        height: 100vh;
        width: calc(45vw);
    }
    .wright {
        justify-content: right;
    }
    .floatLeft {
        float: left;
    }
    .floatRight {
        float: right;
    }
    .marginTop {
        margin-top: 25px;
    }
    .miniWidth {
        width: calc(45vw);
    }
    .fullWidth {
        width: 100%;
    }
    h1 {
        font-size: 3em;
    }
    p {
        font-size: 1.2em;
    }
    li {
        font-size: 1.1em;
    }
    video {
        max-height: 70vh;
        max-width: 35vw;
        border-radius: 8px;
        border: 1px solid var(--text);
    }
    main {
        background: radial-gradient(circle, #4a4d3d, #2f2f2f);
        backdrop-filter: blur(10px);
        width: calc(100vw - 30px);
        height: calc(100vh - 30px);
        padding: 15px;
    }
    .chip {
        width: fit-content;
        cursor: pointer;
    }
</style>
