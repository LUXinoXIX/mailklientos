// Jednoduchý formátovaný editor (TipTap) pre popis udalosti: písmo, farby, zoznamy, odkazy – bez obrázkov a tabuliek.
import { EditorContent, useEditor } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import { Color, FontFamily, FontSize, TextStyle } from '@tiptap/extension-text-style';
import { Highlight } from '@tiptap/extension-highlight';
import { TextAlign } from '@tiptap/extension-text-align';
import { Placeholder } from '@tiptap/extensions';
import { useEffect, useRef } from 'react';
import { FormatBar } from './Composer';

export function RichTextField(props: { html: string; placeholder?: string; onChange: (html: string, text: string) => void }) {
  const onChange = useRef(props.onChange);
  onChange.current = props.onChange;
  const editor = useEditor({
    extensions: [
      StarterKit.configure({ link: { openOnClick: false, autolink: true, defaultProtocol: 'https' } }),
      TextStyle,
      Color,
      FontFamily,
      FontSize,
      Highlight.configure({ multicolor: true }),
      TextAlign.configure({ types: ['heading', 'paragraph'] }),
      Placeholder.configure({ placeholder: props.placeholder ?? '' }),
    ],
    content: props.html,
    shouldRerenderOnTransaction: true,
    editorProps: { attributes: { class: 'composer-body rich-field-body', spellcheck: 'true' } },
    onUpdate: ({ editor: ed }) => {
      const text = ed.getText({ blockSeparator: '\n' }).trim();
      onChange.current(text ? ed.getHTML() : '', text);
    },
  });
  useEffect(() => () => editor?.destroy(), [editor]);
  return (
    <div className="rich-field">
      {editor && <FormatBar editor={editor} table={false} />}
      <EditorContent editor={editor} />
    </div>
  );
}
