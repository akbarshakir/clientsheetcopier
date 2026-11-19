
CREATE TABLE IF NOT EXISTS public.document_copies (
    id BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    document_name TEXT NOT NULL,
    original_url TEXT,
    copied_url TEXT NOT NULL,
    owner_email TEXT NOT NULL,
    replacements_count INTEGER DEFAULT 0,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

ALTER TABLE public.document_copies
ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can read document history"
ON public.document_copies
FOR SELECT
TO authenticated
USING (auth.uid() IS NOT NULL);

CREATE POLICY "Authenticated users can insert document history"
ON public.document_copies
FOR INSERT
TO authenticated
WITH CHECK (auth.uid() IS NOT NULL);
