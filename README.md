# Word-Add-in Static Site (Estate Clause Helper)

This folder contains the production build of the **Estate Clause Helper** Word add-in.  
Publishing these files with **GitHub Pages** makes them publicly reachable so the Word add-in manifest can load its HTML, JS, and image assets.

---

## 1 · Enable GitHub Pages

1. Push this folder (and all its files) to your repository.
2. In **GitHub → Settings → Pages**  
   • **Source** – choose the **`main` branch** and **`/static-site` folder** (or select the `gh-pages` branch if you publish there).  
   • **Build & deployment** – leave as **Static files**.
3. Save. GitHub will deploy and display your live site URL, typically:  
   ```
   https://<username>.github.io/<repo-name>
   ```
4. Wait ~1 minute for the first deploy to finish.

---

## 2 · Update the Add-in Manifest

Copy the URL shown in *Settings → Pages* and replace every occurrence of the old Azure URL in:

* `manifest.xml`
* `taskpane.html` / `taskpane.js`
* `host/*.html` / `host/*.js`

Example change:

```xml
<SourceLocation DefaultValue="https://<username>.github.io/<repo-name>/taskpane.html"/>
```

---

## 3 · Folder Structure Expected by Office

```
static-site/
│  manifest.xml
│  taskpane.html
│  taskpane.js
│
└─host/
   │  function-file.html
   │  function-file.js
   │  taskpane.html
   │  taskpane.js
   │  help.html
   └─assets/
       icon-32.png
       icon.svg
```

---

## 4 · Troubleshooting

| Symptom | Likely Cause | Fix |
|---------|--------------|-----|
| 404 page | Pages still deploying | Wait a minute or check Actions tab |
| Blank add-in pane | Manifest/HTML still points to old URL | Search & replace URLs, then re-sideline the add-in |
| Wrong content at live URL | Wrong branch/folder chosen in Pages settings | Select the folder that contains **this** README |

---

## 5 · Why a README?

Having a README inside the published folder:

* Documents the deployment steps for future updates.  
* Prevents accidental deletion of required files.  
* Makes it obvious which branch/folder is published when browsing the repo.

Push this README along with your static site files, enable Pages, copy the live URL, and your Word add-in will load correctly.
- Walk you through the Admin Center upload and verify the add-in as a test user.

---
Prepared to help with whichever step you want automated next.
