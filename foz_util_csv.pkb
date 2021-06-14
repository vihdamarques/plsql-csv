create or replace package body foz_util_csv is

  function get_pos(p_matriz_csv in t_matriz_csv
                  ,p_linha      in pls_integer
                  ,p_coluna     in pls_integer) return varchar2 is
  begin
    if p_matriz_csv.exists(p_linha) then
      if p_matriz_csv(p_linha).exists(p_coluna) then
        return p_matriz_csv(p_linha)(p_coluna);
      else
        return null;
      end if;
    else
      return null;
    end if;
  end get_pos;

  function a2t(arr in a_array) return t_array is
    l_array t_array;
  begin
    for i in 1 .. arr.count loop
      l_array(i) := arr(i);
    end loop;

    return l_array;
  end a2t;

  function a2t(arr in a_matriz_csv) return t_matriz_csv is
    l_matriz_csv t_matriz_csv;
  begin
    for i in 1 .. arr.count loop
      for j in 1 .. arr(i).count loop
        l_matriz_csv(i)(j) := arr(i)(j);
      end loop;
    end loop;

    return l_matriz_csv;
  end a2t;

  procedure write_clob(p_clob   in out clob
                      ,p_string in varchar2) is
  begin
    dbms_lob.writeappend(p_clob, length(p_string), p_string);
  end;

  function blob_to_clob(p_blob in blob,
                        p_charset in varchar2 default 'WE8ISO8859P1') return clob is
    l_clob         clob;
    l_clob_offset  number := 1;
    l_blob_offset  number := 1;
    l_lang_context number := dbms_lob.default_lang_ctx;
    l_warning      pls_integer;
    --l_pos          number := 1;
    --l_lob_len      number := nvl(dbms_lob.getlength(p_blob), 0);
    --l_buff         varchar2(16000);
  begin
    dbms_lob.createtemporary(l_clob, true);
    dbms_lob.open(l_clob, dbms_lob.lob_readwrite);

    dbms_lob.converttoclob(l_clob
                          ,p_blob
                          ,dbms_lob.lobmaxsize
                          ,l_clob_offset
                          ,l_blob_offset
                          ,nls_charset_id(p_charset) -- AL32UTF8 / WE8ISO8859P1 / dbms_lob.default_csid
                          ,l_lang_context
                          ,l_warning);

    /*loop
      l_buff := utl_raw.cast_to_varchar2(dbms_lob.substr(p_blob
                                                        ,16000
                                                        ,l_pos));
      if length(l_buff) > 0 then
        dbms_lob.writeappend(l_clob, length(l_buff), l_buff);
      end if;

      l_pos := l_pos + 16000;
      exit when l_pos > l_lob_len;
    end loop;*/

    return l_clob; -- l_clob is OPEN here
  end;

  function clob_to_blob(p_clob in clob,
                        p_charset in varchar2 default 'WE8ISO8859P1') return blob is
    l_blob         blob;
    l_clob_offset  number := 1;
    l_blob_offset  number := 1;
    l_lang_context number := dbms_lob.default_lang_ctx;
    l_warning      pls_integer;
    --l_pos     number := 1;
    --l_lob_len number := nvl(dbms_lob.getlength(p_clob), 0);
    --l_amount  number := 16000;
    --l_buff    raw(16000);
  begin
    dbms_lob.createtemporary(l_blob, true);
    dbms_lob.open(l_blob, dbms_lob.lob_readwrite);

    dbms_lob.converttoblob (
      dest_lob      => l_blob,
      src_clob      => p_clob,
      amount        => dbms_lob.lobmaxsize,
      dest_offset   => l_blob_offset,
      src_offset    => l_clob_offset, 
      blob_csid     => nls_charset_id(p_charset), -- AL32UTF8 / WE8ISO8859P1 / dbms_lob.default_csid
      lang_context  => l_lang_context,
      warning       => l_warning
    );

    /*loop
      l_buff := utl_raw.cast_to_raw(dbms_lob.substr(p_clob
                                                   ,l_amount
                                                   ,l_pos));
      if utl_raw.length(l_buff) > 0 then
        dbms_lob.writeappend(l_blob, utl_raw.length(l_buff), l_buff);
      end if;

      l_pos := l_pos + l_amount;
      exit when l_pos > l_lob_len;
    end loop;*/

    return l_blob; -- l_blob is OPEN here
  end;

  procedure download_blob(p_blob        in out blob
                         ,p_file_name   in varchar2
                         ,p_mime_type   in varchar2 default 'application/csv'
                         ,p_disposition in varchar2 default 'attachment') is
  begin
    -- seta header de sessão
    owa_util.mime_header(p_mime_type, false);
    htp.p('Content-length: ' || nvl(dbms_lob.getlength(p_blob), 0));
    htp.p('Content-Disposition: ' || p_disposition ||
          '; filename="' || p_file_name || '"');
    owa_util.http_header_close;

    -- faz download
    sys.wpg_docload.download_file(p_blob);
    apex_application.stop_apex_engine;
  end;

  procedure ler_csv(p_csv            in clob
                   ,p_matriz_csv    out t_matriz_csv
                   ,p_mostrar_titulo in boolean default false) is
    --
    l_clob           clob;
    l_char           char(1);
    l_lookahead      char(1);
    l_pos            number := 0;
    l_token          varchar2(32767) := null;
    l_token_complete boolean := false;
    l_line_complete  boolean := false;
    l_new_token      boolean := true;
    l_enclosed       boolean := false;
    --
    l_last_col_line  boolean := false;
    --
    l_lineno         number := 1;
    l_columnno       number := 1;
    --
    l_columns        t_array;
  begin

    l_clob := p_csv;

    loop
      -- increment position index
      l_pos := l_pos + 1;

      -- get next character from clob
      l_char := dbms_lob.substr( l_clob, 1, l_pos );

      -- exit when no more characters to process
      exit when l_char is null or l_pos > dbms_lob.getLength( l_clob );

      -- if first character of new token is optionally enclosed character
      -- note that and skip it and get next character
      if l_new_token and l_char = G_ENCLOSE then
        l_enclosed := true;
        l_pos := l_pos + 1;
        l_char := dbms_lob.substr( l_clob, 1, l_pos );
      end if;
      l_new_token := false;

      -- get look ahead character
      l_lookahead := dbms_lob.substr( l_clob, 1, l_pos + 1 );

      -- inspect character (and lookahead) to determine what to do
      if l_char = G_ENCLOSE and l_enclosed then

        if l_lookahead = G_ENCLOSE then
          l_pos := l_pos + 1;
          l_token := l_token || l_lookahead;
        elsif l_lookahead = G_DELIM then
          l_pos := l_pos + 1;
          l_token_complete := true;
        else
          l_enclosed := false;
        end if;

      elsif l_char in ( G_CARRIAGE_RETURN, G_LINE_FEED ) and NOT l_enclosed then
        l_token_complete := true;
        l_line_complete := true;

        if l_lookahead in ( G_CARRIAGE_RETURN, G_LINE_FEED ) then
          l_pos := l_pos + 1;
        end if;

      elsif l_char = G_DELIM and not l_enclosed then
        l_token_complete := true;

        if l_pos = dbms_lob.getLength( l_clob ) then -- complete line even if last column is null
          l_line_complete := true;
          l_last_col_line := true;
        end if;

      elsif l_pos = dbms_lob.getLength( l_clob ) then
        l_token          := l_token || l_char;
        l_token_complete := true;
        l_line_complete  := true;

      else
        l_token := l_token || l_char;
      end if;

      -- process a new token
      if l_token_complete then
        l_columns(l_columnno) := l_token;

        l_columnno := l_columnno + 1;
        l_token := null;
        l_enclosed := false;
        l_new_token := true;
        l_token_complete := false;

        if l_last_col_line then -- includes last column even if last column is null
          l_columns(l_columnno) := null;
        end if;
      end if;

      -- process end-of-line here
      if l_line_complete then
        if (l_lineno > 1 or p_mostrar_titulo) then
          p_matriz_csv(l_lineno - case when p_mostrar_titulo then 0 else 1 end) := l_columns;
        end if;

        l_columns.delete;
        l_lineno := l_lineno + 1;
        l_columnno := 1;
        l_line_complete := false;
      end if;
    end loop;
  end ler_csv;

  procedure ler_csv(p_csv            in blob
                   ,p_matriz_csv    out t_matriz_csv
                   ,p_mostrar_titulo in boolean default false
                   ,p_charset        in varchar2 default 'WE8ISO8859P1') is
    --
    l_clob           clob;
    l_blob           blob;
    l_char           char(1);
    l_lookahead      char(1);
    l_pos            number := 0;
    l_token          varchar2(32767) := null;
    l_token_complete boolean := false;
    l_line_complete  boolean := false;
    l_new_token      boolean := true;
    l_enclosed       boolean := false;
    --
    l_lineno         number := 1;
    l_columnno       number := 1;
    --
    l_columns        t_array;
  begin

    l_blob := p_csv;
    l_clob := blob_to_clob(l_blob, p_charset);

    loop
      -- increment position index
      l_pos := l_pos + 1;

      -- get next character from clob
      l_char := dbms_lob.substr( l_clob, 1, l_pos );

      -- exit when no more characters to process
      exit when l_char is null or l_pos > dbms_lob.getLength( l_clob );

      -- if first character of new token is optionally enclosed character
      -- note that and skip it and get next character
      if l_new_token and l_char = G_ENCLOSE then
        l_enclosed := true;
        l_pos := l_pos + 1;
        l_char := dbms_lob.substr( l_clob, 1, l_pos );
      end if;
      l_new_token := false;

      -- get look ahead character
      l_lookahead := dbms_lob.substr( l_clob, 1, l_pos + 1 );

      -- inspect character (and lookahead) to determine what to do
      if l_char = G_ENCLOSE and l_enclosed then

        if l_lookahead = G_ENCLOSE then
          l_pos := l_pos + 1;
          l_token := l_token || l_lookahead;
        elsif l_lookahead = G_DELIM then
          l_pos := l_pos + 1;
          l_token_complete := true;
        else
          l_enclosed := false;
        end if;

      elsif l_char in ( G_CARRIAGE_RETURN, G_LINE_FEED ) and NOT l_enclosed then
        l_token_complete := true;
        l_line_complete := true;

        if l_lookahead in ( G_CARRIAGE_RETURN, G_LINE_FEED ) then
          l_pos := l_pos + 1;
        end if;

      elsif l_char = G_DELIM and not l_enclosed then
        l_token_complete := true;
        
        if l_pos = dbms_lob.getLength( l_clob ) then -- includes last line even if last column is null
          l_line_complete := true;
        end if;

      elsif l_pos = dbms_lob.getLength( l_clob ) then
        l_token := l_token || l_char;
        l_token_complete := true;
        l_line_complete := true;

      else
        l_token := l_token || l_char;
      end if;

      -- process a new token
      if l_token_complete then
        l_columns(l_columnno) := l_token;

        l_columnno := l_columnno + 1;
        l_token := null;
        l_enclosed := false;
        l_new_token := true;
        l_token_complete := false;
        
        if l_pos = dbms_lob.getLength( l_clob ) then -- includes last column even if last column is null
          l_columns(l_columnno) := null;
        end if;
      end if;

      -- process end-of-line here
      if l_line_complete then
        if (l_lineno > 1 or p_mostrar_titulo) then
          p_matriz_csv(l_lineno - case when p_mostrar_titulo then 0 else 1 end) := l_columns;
        end if;

        l_columns.delete;
        l_lineno := l_lineno + 1;
        l_columnno := 1;
        l_line_complete := false;
      end if;
    end loop;
  end ler_csv;

  procedure escrever_csv(p_matriz_csv  in t_matriz_csv
                        ,p_csv        out blob
                        ,p_charset     in varchar2 default 'WE8ISO8859P1') is
    l_column varchar2(32767);
    l_line   clob;
    l_file   clob;
  begin
    dbms_lob.createtemporary(l_line, true);
    dbms_lob.open(l_line, dbms_lob.lob_readwrite);

    dbms_lob.createtemporary(l_file, true);
    dbms_lob.open(l_file, dbms_lob.lob_readwrite);

    for i in 1 .. p_matriz_csv.count() loop
      dbms_lob.trim(l_line, 0);

      for j in 1 .. p_matriz_csv(i).count() loop
        l_column := p_matriz_csv(i)(j);

        if instr(l_column, G_ENCLOSE) > 0 then
          l_column := replace(l_column, G_ENCLOSE, G_ENCLOSE || G_ENCLOSE);
        end if;

        if regexp_instr(l_column, G_DELIM || '|' ||
                                  G_ENCLOSE || '|' ||
                                  G_LINE_FEED || '|' ||
                                  G_CARRIAGE_RETURN) > 0 then
          l_column := G_ENCLOSE || l_column || G_ENCLOSE;
        end if;

        l_column := l_column || G_DELIM;
        dbms_lob.writeappend(l_line, length(l_column), l_column);

      end loop;

      l_line := substr(l_line, 1, nvl(dbms_lob.getlength(l_line), 0) - 1);
      dbms_lob.append(l_file, l_line);
      dbms_lob.writeappend(l_file, length(G_CRLF), G_CRLF);
    end loop;

    p_csv := clob_to_blob(l_file, p_charset);
    dbms_lob.close(l_line);
    dbms_lob.close(l_file);
    dbms_lob.freetemporary(l_line);
    dbms_lob.freetemporary(l_file);
  end escrever_csv;

  procedure escrever_xls(p_matriz_csv  in t_matriz_csv
                        ,p_csv        out blob
                        ,p_charset     in varchar2 default 'WE8ISO8859P1') is
    l_column varchar2(32767);
    l_line   clob;
    l_file   clob;
  begin
    dbms_lob.createtemporary(l_line, true);
    dbms_lob.open(l_line, dbms_lob.lob_readwrite);

    dbms_lob.createtemporary(l_file, true);
    dbms_lob.open(l_file, dbms_lob.lob_readwrite);

    write_clob(l_file, '<?xml version="1.0"?>');
    write_clob(l_file, '<ss:Workbook xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">');

    write_clob(l_file, '<ss:Styles>');
    write_clob(l_file, '<ss:Style ss:ID="OracleDate">');
    write_clob(l_file, '<ss:NumberFormat ss:Format="dd/mm/yyyy\ hh:mm:ss"/>');
    write_clob(l_file, '</ss:Style>');
    write_clob(l_file, '</ss:Styles>');

    write_clob(l_file, '<ss:Worksheet ss:Name="Plan1">');
    write_clob(l_file, '<ss:Table>');

    for i in 1 .. p_matriz_csv.count() loop
      dbms_lob.trim(l_line, 0);

      write_clob(l_line, '<ss:Row>');
      for j in 1 .. p_matriz_csv(i).count() loop
        l_column := p_matriz_csv(i)(j);
        write_clob(l_line, '<ss:Cell>');
        write_clob(l_line, '<ss:Data ss:Type="String">' || l_column || '</ss:Data>');
        write_clob(l_line, '</ss:Cell>');
      end loop;
      write_clob(l_line, '</ss:Row>');
      dbms_lob.append(l_file, l_line);
    end loop;
    write_clob(l_file, '</ss:Table>');
    write_clob(l_file, '</ss:Worksheet>');
    write_clob(l_file, '</ss:Workbook>');

    p_csv := clob_to_blob(l_file, p_charset);
    dbms_lob.close(l_line);
    dbms_lob.close(l_file);
    dbms_lob.freetemporary(l_line);
    dbms_lob.freetemporary(l_file);
  end escrever_xls;

end foz_util_csv;
/