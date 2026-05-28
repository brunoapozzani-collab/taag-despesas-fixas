# TAAG Reference CSS — Ground Truth

Source: https://taagsystemoficial.onrender.com/  
Captured: 2026-05-05  
Method: `curl -s -L https://taagsystemoficial.onrender.com/ -o taag-reference.html`  
External CSS files: **none** — all styles are inline in a single `<style>` block.  
Font CDN: `https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600&display=swap`

---

## Canonical iOS Safari Background Pattern

`background-attachment: fixed` on `body` is broken on iOS Safari — renders as `scroll`.  
**Always use this pattern instead:**

```css
body::before {
  content: "";
  position: fixed;
  inset: 0;
  z-index: -1;
  background-image: image-set(
    url('/bg.webp') type('image/webp'),
    url('/bg.jpg')  type('image/jpeg')
  );
  background-size: cover;
  background-position: center;
}
body { background: #0e0e0e; }  /* fallback during image load */
```

Works on: iOS Safari, Android Chrome, all desktop browsers.  
The `image-set()` serves WebP to supporting browsers and JPEG as fallback.  
Images: `/public/bg.webp` (119 KB) and `/public/bg.jpg` (199 KB), both 1920×1440.

---

## Ground-Truth Corrections (from pre-flight Check 3)

| Property | Plan spec (wrong) | Actual value (correct) |
|---|---|---|
| Input border | `rgba(255,255,255,0.10)` | `rgba(255,255,255,0.20)` |
| Input focus border | not specified | `rgba(255,255,255,0.50)` |
| Card border | `rgba(255,255,255,0.10)` | `rgba(255,255,255,0.10)` ✅ unchanged |

**This file wins over any earlier spec when values conflict.**

---

## Property Reference Table

| Property | Value | Selector | Notes |
|---|---|---|---|
| **Page background color** | `#0e0e0e` | `body` | Near-black fallback behind image |
| **Background image URL** | `/static/assets/img/WhatsApp%20Image%202025-09-26%20at%2011.48.34.jpeg` | `body` | Dark office interior, 4032×3024px, 624KB |
| **Background size** | `cover` | `body` | |
| **Background position** | `center` | `body` | |
| **Background repeat** | `no-repeat` | `body` | |
| **Background attachment** | `fixed` | `body` | Parallax-like; critical for glassmorphism feel |
| **Body layout** | `display: flex; align-items: center; justify-content: center` | `body` | Full-screen centering |
| **Body overflow** | `hidden` | `body` | Overridden to `auto` on mobile |
| **Font family** | `'Poppins', sans-serif` | `*` | Applied universally via reset |
| **Font weights loaded** | `300, 400, 600` | Google Fonts | |
| **Card max-width** | `800px` | `.container` | |
| **Card min-height** | `450px` | `.container` | |
| **Card width** | `90%` | `.container` | Responsive shrink |
| **Card border-radius** | `20px` | `.container` | |
| **Card background** | `rgba(255, 255, 255, 0.05)` | `.container` | Ultra-translucent glass base |
| **Card backdrop-filter** | `blur(15px)` | `.container` | Both prefixed and unprefixed |
| **Card -webkit-backdrop-filter** | `blur(15px)` | `.container` | Required for Safari |
| **Card border** | `1px solid rgba(255, 255, 255, 0.1)` | `.container` | Hairline white |
| **Card box-shadow** | `0 8px 32px rgba(0, 0, 0, 0.7)` | `.container` | Deep dark shadow |
| **Card overflow** | `hidden` | `.container` | Clips child panels to border-radius |
| **Card layout** | `display: flex; flex-direction: row` | `.container` | Side-by-side panels |
| **Left panel background** | `rgba(255, 255, 255, 0.1)` | `.left` | Slightly more opaque than card |
| **Left panel backdrop-filter** | `blur(15px)` + `-webkit-backdrop-filter: blur(15px)` | `.left` | Both required |
| **Left panel padding** | `40px` | `.left` | |
| **Logo max-width** | `250px` | `.left img` | |
| **Logo border-radius** | `40%` | `.left img` | Softly rounds the corners |
| **Right panel background** | `rgba(255, 255, 255, 0.03)` | `.right` | Near-invisible glass |
| **Right panel padding** | `40px` | `.right` | |
| **Right panel border-left** | `1px solid rgba(255, 255, 255, 0.1)` | `.right` | Separator hairline |
| **Heading color** | `#fff` | `.right h2` | |
| **Heading font-weight** | `600` | `.right h2` | |
| **Heading letter-spacing** | `1px` | `.right h2` | |
| **Heading margin-bottom** | `30px` | `.right h2` | |
| **Input padding** | `14px 20px` | `.form-control` | |
| **Input border-radius** | `30px` | `.form-control` | Pill shape |
| **Input border** | `1px solid rgba(255, 255, 255, 0.2)` | `.form-control` | Slightly brighter than card border |
| **Input background** | `rgba(255, 255, 255, 0.1)` | `.form-control` | |
| **Input color** | `#fff` | `.form-control` | |
| **Input font-size** | `16px` | `.form-control` | |
| **Input transition** | `0.3s` | `.form-control` | All properties |
| **Input placeholder color** | `rgba(255, 255, 255, 0.5)` | `.form-control::placeholder` | Half-opacity white |
| **Input focus background** | `rgba(255, 255, 255, 0.2)` | `.form-control:focus` | Brightens on focus |
| **Input focus border-color** | `rgba(255, 255, 255, 0.5)` | `.form-control:focus` | Brighter hairline |
| **Input error border-color** | `#ff6b6b` | `.form-control.input-error` | |
| **Input error background** | `rgba(255, 107, 107, 0.12)` | `.form-control.input-error` | |
| **Password toggle color** | `rgba(255, 255, 255, 0.5)` | `.toggle-pw` | |
| **Password toggle hover** | `#fff` | `.toggle-pw:hover` | |
| **CapsLock warning bg** | `rgba(255, 204, 0, 0.15)` | `.capslock-warn` | |
| **CapsLock warning border** | `1px solid rgba(255, 204, 0, 0.4)` | `.capslock-warn` | |
| **CapsLock warning color** | `#ffd60a` | `.capslock-warn` | Bright yellow |
| **CapsLock warning border-radius** | `20px` | `.capslock-warn` | Pill |
| **Alert border-radius** | `12px` | `.alert` | Rounded rectangle (not pill) |
| **Alert error color** | `#ff6b6b` | `.alert-error` | |
| **Alert error background** | `rgba(255, 107, 107, 0.12)` | `.alert-error` | |
| **Alert error border** | `1px solid rgba(255, 107, 107, 0.3)` | `.alert-error` | |
| **Alert warning color** | `#ffd60a` | `.alert-warning` | |
| **Alert warning background** | `rgba(255, 214, 10, 0.1)` | `.alert-warning` | |
| **Alert warning border** | `1px solid rgba(255, 214, 10, 0.3)` | `.alert-warning` | |
| **Button padding** | `14px` | `.btn` | Vertical only; full width |
| **Button border-radius** | `30px` | `.btn` | Pill shape |
| **Button background** | `#fff` | `.btn` | Pure white fill |
| **Button color** | `#000` | `.btn` | Black text |
| **Button font-size** | `16px` | `.btn` | |
| **Button font-weight** | `600` | `.btn` | Semi-bold |
| **Button transition** | `0.3s` | `.btn` | |
| **Button margin-top** | `4px` | `.btn` | |
| **Button hover background** | `#ddd` | `.btn:hover` | Light grey |
| **Button hover transform** | `translateY(-2px)` | `.btn:hover` | Lifts up |
| **Footer link color** | `#ccc` | `.footer-links a` | Muted grey |
| **Footer link hover** | `#fff` | `.footer-links a:hover` | |
| **Footer link font-size** | `14px` | `.footer-links a` | |
| **Mobile breakpoint** | `max-width: 768px` | `@media` | |
| **Mobile card direction** | `column` | `.container` (mobile) | Stacks vertically |
| **Mobile card max-width** | `380px` | `.container` (mobile) | |
| **Mobile right border** | `border-top: 1px solid rgba(255,255,255,0.1)` | `.right` (mobile) | Becomes top border, not left |

---

## Background Image

| Property | Value |
|---|---|
| Path on server | `/static/assets/img/WhatsApp%20Image%202025-09-26%20at%2011.48.34.jpeg` |
| Downloaded to | `/tmp/taag-bg-original.jpg` |
| Dimensions | 4032 × 3024 px |
| File size | 639,002 bytes (≈ 624 KB) |
| Format | JPEG progressive, JFIF 1.01 |
| Content | Dark office interior — wooden conference table, tropical plants lit from behind, dark ceiling, shelving. Very dark overall, ideal for glassmorphism overlay. |

---

## Notes for Next.js Implementation

1. All CSS is inline — no external stylesheet to fetch or bundle. In Next.js, translate to Tailwind + CSS variables in `globals.css` or a dedicated `taag.css` module.
2. `background-attachment: fixed` may render incorrectly on iOS Safari (known bug — `fixed` background doesn't scroll correctly inside elements with `overflow`). Workaround: use a `position: fixed` pseudo-element for the background layer instead of setting it on `body`.
3. Both `backdrop-filter` and `-webkit-backdrop-filter` must be present — Safari requires the prefixed version.
4. No JavaScript libraries are used — all interactivity (password toggle, CapsLock detection) is vanilla JS. Translate to React state/event handlers.
5. The font CDN call is `Poppins:wght@300;400;600` — load in Next.js via `next/font/google` for performance (automatic preload, zero layout shift).
