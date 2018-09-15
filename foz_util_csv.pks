create or replace package foz_util_csv is

  G_DELIM           char(1) := ';';
  G_ENCLOSE         char(1) := '"';
  G_CARRIAGE_RETURN constant char(1) := chr(13);
  G_LINE_FEED       constant char(1) := chr(10);
  G_CRLF            char(2) := G_CARRIAGE_RETURN || G_LINE_FEED;

  -- Usar para inicialização do tipo: l_array[1] = 1; l_array[2] = 2;
  type t_array      is table of varchar2(32767) index by pls_integer;
  type t_matriz_csv is table of t_array index by pls_integer;

  -- Usar para inicialização do tipo: l_array := a_array(1,2,3)
  type a_array      is table of varchar2(32767);
  type a_matriz_csv is table of a_array;

  -- Usar para converter variaveis do tipo a_ (array)
  -- para o tipo t_ (Array Associativo)
  function a2t(arr in a_array) return t_array;
  function a2t(arr in a_matriz_csv) return t_matriz_csv;

  -- Utilitarios
  function blob_to_clob(p_blob in blob) return clob;
  function clob_to_blob(p_clob in clob) return blob;
  procedure download_blob(p_blob        in out blob
                         ,p_file_name   in varchar2
                         ,p_mime_type   in varchar2 default 'application/csv'
                         ,p_disposition in varchar2 default 'attachment');

  -- Pegar informação por posição
  function get_pos(p_matriz_csv in t_matriz_csv
                  ,p_linha      in pls_integer
                  ,p_coluna     in pls_integer) return varchar2;

  -- Ler arquivo CSV
/*
  Exemplo de leitura:
  declare
    v_matriz_csv foz_util_csv.t_matriz_csv;
    v_blob blob;
  begin
    select ffsa_bin_arquivo
      into v_blob
      from FOZ_FRM_SOLICITACAO_ANEXOS
     where ffsa_id_frm_anexo = 49328;

    foz_util_csv.ler_csv(v_blob
                        ,v_matriz_csv);

    for i in 1 .. v_matriz_csv.count loop
      for j in 1 .. v_matriz_csv(i).count loop
        dbms_output.put(v_matriz_csv(i)(j) || ' | ');
      end loop;
      dbms_output.put_line(null);
    end loop;
  end;
*/
  procedure ler_csv(p_csv            in blob
                   ,p_matriz_csv    out t_matriz_csv
                   ,p_mostrar_titulo in boolean default false);

  -- Escrever arquivo CSV
/*
  Teste de escrita:
  declare
    l_matriz_csv foz_util_csv.t_matriz_csv;
    l_csv blob;
  begin
    l_matriz_csv(1)(1) := 'Teste';
    l_matriz_csv(1)(2) := 'Tes$#t"e';
    l_matriz_csv(1)(3) := 'he
  h';
    l_matriz_csv(2)(1) := 'máis uma linha';
    l_matriz_csv(2)(2) := 'AB;D';
    l_matriz_csv(2)(3) := '21ç65';

    foz_util_csv.escrever_csv(l_matriz_csv, l_csv);
    dbms_output.put_line(foz_util_csv.blob_to_clob(l_csv));
  end;

  ou

  declare
    l_csv foz_util_csv.t_matriz_csv;
    l_csv_blob blob;
  begin
    l_csv := foz_util_csv.a2t(foz_util_csv.a_matriz_csv(
                 foz_util_csv.a_array('Teste', 'Tes$#t"e', 'heh'),
                 foz_util_csv.a_array('máis uma linha', 'AB;D', '21ç65')
             ));
    l_csv(l_csv.count+1) := foz_util_csv.a2t(foz_util_csv.a_array('x', 'y', 'z'));
    foz_util_csv.escrever_csv(l_csv, l_csv_blob);
    dbms_output.put_line(foz_util_csv.blob_to_clob(l_csv_blob));
  end;

  ou

  declare
    l_csv foz_util_csv.t_matriz_csv;
    l_csv_blob blob;
  begin
    l_csv(l_csv.count + 1) := foz_util_csv.a2t(foz_util_csv.a_array('x', 'y', 'z'));
    l_csv(l_csv.count + 1) := foz_util_csv.a2t(foz_util_csv.a_array('a', 'b', 'c'));
    foz_util_csv.escrever_csv(l_csv, l_csv_blob);
    dbms_output.put_line(foz_util_csv.blob_to_clob(l_csv_blob));
  end;
*/
  procedure escrever_csv(p_matriz_csv  in t_matriz_csv
                        ,p_csv        out blob);
  -- XML Excel  *BETA*
  procedure escrever_xls(p_matriz_csv  in t_matriz_csv
                        ,p_csv        out blob);
end foz_util_csv;