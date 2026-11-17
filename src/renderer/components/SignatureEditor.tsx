// Formátovaný editor podpisu (TipTap): písmo, veľkosť, farby, odkazy, logo/obrázky so zmenou veľkosti, tabuľka na rozloženie.
import { EditorContent, useEditor } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import { Color, FontFamily, FontSize, TextStyle } from '@tiptap/extension-text-style';
import { Highlight } from '@tiptap/extension-highlight';
import { TextAlign } from '@tiptap/extension-text-align';
import { Image } from '@tiptap/extension-image';
import { TableKit } from '@tiptap/extension-table';
import { Placeholder } from '@tiptap/extensions';
import { useEffect, useRef } from 'react';
import { FormatBar, readAsDataUrl } from './Composer';

/** Logo do podpisu: zmenšiť veľké obrázky (max. 600 px na šírku), aby podpis nemal megabajty. */
async function shrinkImage(f: File): Promise<{ src: string; width?: number }> {
  const src = await readAsDataUrl(f);
  const img = new window.Image();
  await new Promise((res, rej) => {
    img.onload = res;
    img.onerror = rej;
    img.src = src;
  });
  // zobrazená šírka v podpise: logo zvyčajne do 200 px (dá sa zmeniť ťahaním za roh)
  const width = Math.min(img.naturalWidth || 200, 200);
  const max = 600;
  if (f.type === 'image/svg+xml' || f.type === 'image/gif' || img.naturalWidth <= max) return { src, width };
  const c = document.createElement('canvas');
  c.width = max;
  c.height = Math.round((img.naturalHeight * max) / img.naturalWidth);
  c.getContext('2d')?.drawImage(img, 0, 0, c.width, c.height);
  return { src: c.toDataURL(f.type === 'image/jpeg' ? 'image/jpeg' : 'image/png', 0.9), width };
}

export function SignatureEditor(props: { id: string; html: string; onChange: (html: string, text: string) => void }) {
  const onChange = useRef(props.onChange);
  onChange.current = props.onChange;
  const fileInput = useRef<HTMLInputElement>(null);

  const insertFiles = (files: File[]) => {
    const imgs = files.filter((f) => f.type.startsWith('image/'));
    void Promise.all(imgs.map(shrinkImage)).then((srcs) => srcs.forEach((im) => editor?.chain().focus().setImage(im).run()));
    return imgs.length > 0;
  };

  const editor = useEditor(
    {
      extensions: [
        StarterKit.configure({ link: { openOnClick: false, autolink: true, defaultProtocol: 'https' } }),
        TextStyle,
        Color,
        FontFamily,
        FontSize,
        Highlight.configure({ multicolor: true }),
        TextAlign.configure({ types: ['heading', 'paragraph'] }),
        Image.configure({ inline: true, allowBase64: true, resize: { enabled: true, alwaysPreserveAspectRatio: true, minWidth: 16, minHeight: 16 } }),
        TableKit.configure({ table: { resizable: false } }),
        Placeholder.configure({ placeholder: 'Napíšte podpis: meno, funkcia, firma, telefón… logo vložíte tlačidlom Obrázok alebo pretiahnutím.' }),
      ],
      content: props.html,
      shouldRerenderOnTransaction: true,
      editorProps: {
        attributes: { class: 'composer-body sig-body', spellcheck: 'true' },
        handlePaste: (_v, e) => insertFiles([...(e.clipboardData?.files ?? [])]),
        handleDrop: (_v, e) => insertFiles([...((e as DragEvent).dataTransfer?.files ?? [])]),
      },
      onUpdate: ({ editor: ed }) => onChange.current(ed.getHTML(), ed.getText({ blockSeparator: '\n' }).trim()),
    },
    [props.id],
  );

  // pri prepnutí na iný podpis editor vznikne nanovo (závislosť props.id); obsah sa nastaví v content
  useEffect(() => () => editor?.destroy(), [editor]);

  return (
    <div className="sig-editor">
      {editor && <FormatBar editor={editor} onImage={() => fileInput.current?.click()} />}
      <div className="sig-editor-body">
        <EditorContent editor={editor} />
      </div>
      <input
        ref={fileInput}
        type="file"
        accept="image/*"
        hidden
        onChange={(e) => {
          insertFiles([...(e.target.files ?? [])]);
          e.target.value = '';
        }}
      />
    </div>
  );
}
