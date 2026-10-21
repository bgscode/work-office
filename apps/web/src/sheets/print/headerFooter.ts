// Excel header/footer codes -> text. A header is up to three sections:
// `&L` left, `&C` centre, `&R` right (text before any section code is centred).
// Field codes: &P page, &N pages, &D date, &T time, &F file, &A sheet, &Z path,
// && a literal ampersand. Font codes (&B, &I, &U, &S, &E, &X, &Y, &"Font,Style",
// &nn size, &Kcolour, &G picture) are recognised and dropped.

export interface HeaderFooterContext {
  page: number;
  pages: number;
  sheet: string;
  file: string;
  date: Date;
}

export interface Sections {
  left: string;
  center: string;
  right: string;
}

export function renderHeaderFooter(code: string, ctx: HeaderFooterContext): Sections {
  const out: Sections = { left: "", center: "", right: "" };
  let section: keyof Sections = "center";
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch !== "&" || i + 1 >= code.length) {
      out[section] += ch;
      continue;
    }
    const next = code[++i];
    // `&P+1` / `&P-1`: page number offsets.
    const readOffset = (): number => {
      const m = /^([+-])(\d+)/.exec(code.slice(i + 1));
      if (!m) return 0;
      i += m[0].length;
      return (m[1] === "-" ? -1 : 1) * Number(m[2]);
    };
    switch (next) {
      case "L":
        section = "left";
        break;
      case "C":
        section = "center";
        break;
      case "R":
        section = "right";
        break;
      case "P":
        out[section] += String(ctx.page + readOffset());
        break;
      case "N":
        out[section] += String(ctx.pages);
        break;
      case "D":
        out[section] += ctx.date.toLocaleDateString();
        break;
      case "T":
        out[section] += ctx.date.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
        break;
      case "F":
        out[section] += ctx.file;
        break;
      case "A":
        out[section] += ctx.sheet;
        break;
      case "Z":
        break;
      case "&":
        out[section] += "&";
        break;
      case '"': {
        // &"Font name,Style"
        const end = code.indexOf('"', i + 1);
        i = end === -1 ? code.length : end;
        break;
      }
      case "K":
        // &KRRGGBB or theme colour &KTTSNN: six characters.
        i += 6;
        break;
      default:
        if (/[0-9]/.test(next)) {
          // Font size: one or more digits.
          while (i + 1 < code.length && /[0-9]/.test(code[i + 1])) i++;
        }
        // B I U S E X Y G and unknown codes: formatting only.
        break;
    }
  }
  return out;
}

export function isEmptySections(s: Sections): boolean {
  return !s.left && !s.center && !s.right;
}
