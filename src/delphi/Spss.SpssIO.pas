unit Spss.SpssIO;

interface
uses Windows, Dialogs;




const LC_ALL    =  0;

const LC_COLLATE = 1;

const LC_CTYPE   = 2;



//****************************************** VARIABLE TYPES
const SPSS_STRING   = 1;
const SPSS_NUMERIC  = 0;

//****************************************** MISSING VALUE TYPE CODES
const SPSS_NO_MISSVAL       =     0;
const SPSS_ONE_MISSVAL      =     1;
const SPSS_TWO_MISSVAL      =     2;
const SPSS_THREE_MISSVAL    =     3;
const SPSS_MISS_RANGE       =    -2;
const SPSS_MISS_RANGEANDVAL =    -3;

//****************************************** ERROR CODES
const SPSS_OK                     =   0;

const SPSS_FILE_OERROR            =   1;
const SPSS_FILE_WERROR            =   2;
const SPSS_FILE_RERROR            =   3;
const SPSS_FITAB_FULL             =   4;
const SPSS_INVALID_HANDLE         =   5;
const SPSS_INVALID_FILE           =   6;
const SPSS_NO_MEMORY              =   7;

const SPSS_OPEN_RDMODE            =   8;
const SPSS_OPEN_WRMODE            =   9;

const SPSS_INVALID_VARNAME        =   10;
const SPSS_DICT_EMPTY             =   11;
const SPSS_VAR_NOTFOUND           =   12;
const SPSS_DUP_VAR                =   13;
const SPSS_NUME_EXP               =   14;
const SPSS_STR_EXP                =   15;
const SPSS_SHORTSTR_EXP           =   16;
const SPSS_INVALID_VARTYPE        =   17;

const SPSS_INVALID_MISSFOR        =   18;
const SPSS_INVALID_COMPSW         =   19;
const SPSS_INVALID_PRFOR          =   20;
const SPSS_INVALID_WRFOR          =   21;
const SPSS_INVALID_DATE           =   22;
const SPSS_INVALID_TIME           =   23;

const SPSS_NO_VARIABLES           =   24;
const SPSS_MIXED_TYPES            =   25;
const SPSS_DUP_VALUE              =   27;

const SPSS_INVALID_CASEWGT        =   28;
const SPSS_INCOMPATIBLE_DICT      =   29;
const SPSS_DICT_COMMIT            =   30;
const SPSS_DICT_NOTCOMMIT         =   31;

const SPSS_NO_TYPE2               =   33;
const SPSS_NO_TYPE73              =   41;
const SPSS_INVALID_DATEINFO       =   45;
const SPSS_NO_TYPE999             =   46;
const SPSS_EXC_STRVALUE           =   47;
const SPSS_CANNOT_FREE            =   48;
const SPSS_BUFFER_SHORT           =   49;
const SPSS_INVALID_CASE           =   50;
const SPSS_INTERNAL_VLABS         =   51;
const SPSS_INCOMPAT_APPEND        =   52;
const SPSS_INTERNAL_D_A           =   53;
const SPSS_FILE_BADTEMP           =   54;
const SPSS_DEW_NOFIRST            =   55;
const SPSS_INVALID_MEASURELEVEL   =   56;
const SPSS_INVALID_7SUBTYPE       =   57;
const SPSS_INVALID_VARHANDLE      =   58;
const SPSS_INVALID_ENCODING       =   59;
const SPSS_FILES_OPEN             =   60;

const SPSS_INVALID_MRSETDEF       =   70;
const SPSS_INVALID_MRSETNAME      =   71;
const SPSS_DUP_MRSETNAME          =   72;

const SPSS_BAD_EXTENSION          =   73;
const SPSS_INVALID_EXTENDEDSTRING =   74;

const SPSS_INVALID_ATTRNAME       =   75;
const SPSS_INVALID_ATTRDEF        =   76;
const SPSS_INVALID_MRSETINDEX     =   77;


//****************************************** WARNING CODES
const SPSS_EXC_LEN64            = -1;
const SPSS_EXC_LEN120           = -2;     //* Retain for compatibility */
const SPSS_EXC_VARLABEL         = -2;
const SPSS_EXC_LEN60            = -4;
const SPSS_EXC_VALLABEL         = -4;
const SPSS_FILE_END             = -5;
const SPSS_NO_VARSETS           = -6;
const SPSS_EMPTY_VARSETS        = -7;
const SPSS_NO_LABELS            = -8;
const SPSS_NO_LABEL             = -9;
const SPSS_NO_CASEWGT           = -10;
const SPSS_NO_DATEINFO          = -11;
const SPSS_NO_MULTRESP          = -12;
const SPSS_EMPTY_MULTRESP       = -13;
const SPSS_NO_DEW               = -14;
const SPSS_EMPTY_DEW            = -15;

//****************************************** FORMAT TYPE CODES
const SPSS_FMT_A          =  1;     //* Alphanumeric */
const SPSS_FMT_AHEX       =  2;     //* Alphanumeric hexadecimal */
const SPSS_FMT_COMMA      =  3;     //* F Format with commas */
const SPSS_FMT_DOLLAR     =  4;     //* Commas and floating dollar sign */
const SPSS_FMT_F          =  5;     //* Default Numeric Format */
const SPSS_FMT_IB         =  6;     //* Integer binary */
const SPSS_FMT_PIBHEX     =  7;     //* Positive integer binary - hex */
const SPSS_FMT_P          =  8;     //* Packed decimal */
const SPSS_FMT_PIB        =  9;     //* Positive integer binary unsigned */
const SPSS_FMT_PK         = 10;     //* Positive integer binary unsigned */
const SPSS_FMT_RB         = 11;     //* Floating point binary */
const SPSS_FMT_RBHEX      = 12;     //* Floating point binary hex */
const SPSS_FMT_Z          = 15;     //* Zoned decimal */
const SPSS_FMT_N          = 16;     //* N Format- unsigned with leading 0s */
const SPSS_FMT_E          = 17;     //* E Format- with explicit power of 10 */
const SPSS_FMT_DATE       = 20;     //* Date format dd-mmm-yyyy */
const SPSS_FMT_TIME       = 21;     //* Time format hh:mm:ss.s  */
const SPSS_FMT_DATE_TIME  = 22;     //* Date and Time           */
const SPSS_FMT_ADATE      = 23;     //* Date format mm/dd/yyyy */
const SPSS_FMT_JDATE      = 24;     //* Julian date - yyyyddd   */
const SPSS_FMT_DTIME      = 25;     //* Date-time dd hh:mm:ss.s */
const SPSS_FMT_WKDAY      = 26;     //* Day of the week         */
const SPSS_FMT_MONTH      = 27;     //* Month                   */
const SPSS_FMT_MOYR       = 28;     //* mmm yyyy                */
const SPSS_FMT_QYR        = 29;     //* q Q yyyy                */
const SPSS_FMT_WKYR       = 30;     //* ww WK yyyy              */
const SPSS_FMT_PCT        = 31;     //* Percent - F followed by % */
const SPSS_FMT_DOT        = 32;     //* Like COMMA, switching dot for comma */
const SPSS_FMT_CCA        = 33;     //* User Programmable currency format   */
const SPSS_FMT_CCB        = 34;     //* User Programmable currency format   */
const SPSS_FMT_CCC        = 35;     //* User Programmable currency format   */
const SPSS_FMT_CCD        = 36;     //* User Programmable currency format   */
const SPSS_FMT_CCE        = 37;     //* User Programmable currency format   */
const SPSS_FMT_EDATE      = 38;     //* Date in dd.mm.yyyy style            */
const SPSS_FMT_SDATE      = 39;     //* Date in yyyy/mm/dd style            */

//****************************************** MEASUREMENT LEVEL CODES
const SPSS_MLVL_UNK       =  0;     //* Unknown */
const SPSS_MLVL_NOM       =  1;     //* Nominal */
const SPSS_MLVL_ORD       =  2;     //* Ordinal */
const SPSS_MLVL_RAT       =  3;     //* Scale(Ratio) */

//****************************************** ALIGNMENT CODES
const SPSS_ALIGN_LEFT     =  0;
const SPSS_ALIGN_RIGHT    =  1;
const SPSS_ALIGN_CENTER   =  2;

//****************************************** DIAGNOSTICS REGARDING VAR NAMES
const SPSS_NAME_OK        = 0;   //* Valid standard name */
const SPSS_NAME_SCRATCH   = 1;   //* Valid scratch var name */
const SPSS_NAME_SYSTEM    = 2;   //* Valid system var name */
const SPSS_NAME_BADLTH    = 3;   //* Empty or longer than SPSS_MAX_VARNAME */
const SPSS_NAME_BADCHAR   = 4;   //* Invalid character or imbedded blank */
const SPSS_NAME_RESERVED  = 5;   //* Name is a reserved word */
const SPSS_NAME_BADFIRST  = 6;   //* Invalid initial character */

//****************************************** MAXIMUM LENGTHS OF DATA FILE OBJECTS
const SPSS_MAX_VARNAME      = 64;     //* Variable name */
const SPSS_MAX_SHORTVARNAME = 8;      //* Short (compatibility) variable name */
const SPSS_MAX_SHORTSTRING  = 8;      // * Short string variable */
const SPSS_MAX_IDSTRING     = 64;     //* File label string */
const SPSS_MAX_LONGSTRING   = 32767;  //* Long string variable */
const SPSS_MAX_VALLABEL     = 120;    //* Value label */
const SPSS_MAX_VARLABEL     = 256;    //* Variable label */
const SPSS_MAX_ENCODING     = 64;     //* Maximum encoding text */
const SPSS_MAX_7SUBTYPE     = 40;     //* Maximum record 7 subtype */

//****************************************** Type 7 record subtypes
const SPSS_T7_DOCUMENTS     = 0;     //* Documents (actually type 6 */
const SPSS_T7_VAXDE_DICT    = 1;     //* VAX Data Entry - dictionary version */
const SPSS_T7_VAXDE_DATA    = 2;     //* VAX Data Entry - data */
const SPSS_T7_SOURCE        = 3;     //* Source system characteristics */
const SPSS_T7_HARDCONST     = 4;     //* Source system floating pt constants */
const SPSS_T7_VARSETS       = 5;     //* Variable sets */
const SPSS_T7_TRENDS        = 6;     //* Trends date information */
const SPSS_T7_MULTRESP      = 7;     //* Multiple response groups */
const SPSS_T7_DEW_DATA      = 8;     //* Windows Data Entry data */
const SPSS_T7_TEXTSMART     = 10;    //* TextSmart data */
const SPSS_T7_MSMTLEVEL     = 11;    //* Msmt level, col width, & alignment */
const SPSS_T7_DEW_GUID      = 12;    //* Windows Data Entry GUID */
const SPSS_T7_XVARNAMES     = 13;    //* Extended variable names */
const SPSS_T7_XSTRINGS      = 14;    //* Extended strings */
const SPSS_T7_CLEMENTINE    = 15;    //* Clementine Metadata */
const SPSS_T7_NCASES        = 16;    //* 64 bit N of cases */
const SPSS_T7_FILE_ATTR     = 17;    //* File level attributes */
const SPSS_T7_VAR_ATTR      = 18;    //* Variable attributes */
const SPSS_T7_EXTMRSETS     = 19;    //* Extended multiple response groups */
const SPSS_T7_ENCODING      = 20;    //* Encoding, aka code page */
const SPSS_T7_LONGSTRLABS   = 21;    //* Value labels for long strings */                                .
const SPSS_T7_LONGSTRMVAL   = 22;    //* Missing values for long strings */

//****************************************** Encoding modes
const SPSS_ENCODING_CODEPAGE   = 0;  //* Text encoded in current code page */
const SPSS_ENCODING_UTF8       = 1;  //* Text encoded as UTF-8 */


type
  {Para PAnsiChar                                }
  {*   - ponteiro de array (PAnsiChar)           }
  {**  - Var de ponteiro de array            }
  {*** - Var de um ponteiro de array de PAnsiChar}

  {Para int e LongInt                                 }
  {*   - Var de int ou longint                        }
  {**  - Ponteiro de array de int ou longint          }
  {*** - Var de um ponteiro de array de int ou longint}

  TValuec = array[0..65535] of PAnsiChar;
  PTValuec = ^TValuec;

  TValuei = array[0..65535] of integer;
  PTValuei = ^TValuei;

  TValued = array[0..65535] of Double;
  PTValued = ^TValued;

  TValuelong = array[0..65535] of Longint;
  PTValuelong = ^TValuelong;

  {Estrutura para leitura e escrita de multi resposta}
  PspssMultRespDef = ^TspssMultRespDef;
  TspssMultRespDef = record
    {Nome p/ multi resposta terminado em null, tamanho máximo de 64 inicia "$" }
    szMrSetName        : array[0..SPSS_MAX_VARNAME] of Char;

    {Label(Pergunta) para multi resposta terminado em null, tamanho de 256     }
    szMrSetLabel       : array[0..SPSS_MAX_VARLABEL] of Char;

    {1 se multi resposta dictioma e 0 se não                                   }
    qIsDichotomy       : integer;

    {1 se contém valor numerico e 0 se não                                     }
    qIsNumeric         : integer;

    {1 se várias categorias para as quais o dicotomias estão a ser marcados com}
    {o valor correspondente rótulos ao valor contabilizado.                    }
    qUseCategoryLabels : integer;

    {1 se para múltiplas dicotomias para a qual o rótulo para o conjunto ?    }
    {retirado do rótulo da variável primeiro componente variável               }
    qUseFirstVarLabel  : integer;

    {Reservado para expansão}
    Reserved           : array[0..13] of integer;

    {O valor contabilizado para numérico múltiplos dicotomias                  }
    nCountedValue      : longint;

    {Ponteiro o valor contabilizado para string de múltiplos dicotomias        }
    pszCountedValue    : PAnsiChar;

    {Ponteiro para o array de nomes}
    ppszVarNames       : PTValuec;

    {Numero de variáveis da estrutura}
    nVariables          : integer;
  end;

  {Manipulação de arquivos}

  {Esta função abre o aruqivo para leitura, retornando handle para operações no}
  {arquivo                                                                     }
  {filename     - Caminho e nome do arquivo que ser?aberto p/ leitura         }
  {handle       - Handle do arquivo.                                           }
  TspssOpenRead              = function(const fileName : PAnsiChar; var handle : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função abre o arquivo abrindo os cases e retornando o handle para      }
  {operações no arquivo                                                        }
  TspssOpenAppend            = function(const fileName : PAnsiChar; var handle : Integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função abre um novo arquivo para escrita                               }
  TspssOpenWrite             = function(const fileName : PAnsiChar; var handle : Integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}


  {Esta função fecha arquivo aberto pela spssOpenRead  somente por esta função }
  {handle       - Handle do arquivo.                                           }
  TspssCloseRead             = function(const handle : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {28/04/2008}

  {Esta função fecha arquivo aberto pela spssOpenAppend somente por esta função}
  {handle       - Handle do arquivo.                                           }
  TspssCloseAppend           = function(const handle : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função fecha arquivo aberto pela spssOpenWrite somente por esta função }
  {handle       - Handle do arquivo.                                           }
  TspssCloseWrite            = function(const handle : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função copia os dados do arquivo de origem para o de destino           }
  {O arquivo de destino tem que ser aberto como Output. Se o arquivo de destino}
  {j?estiver aberto a copia ser?descartada. Se o arquivo de origem não tiver }
  {informações o destino também ficar?vazio                                   }
  {fromHandle     - Handle do arquivo de origem                                }
  {toHandle       - Handle do arquivo de destino                               }
  TspssCopyDocuments         = function(const fromHandle : integer; Const toHandle : Integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}



  {Pega os dados das pesquisas(Pergunta, repostas)}
  TspssGetVarNValueLabels    = function(const handle : integer; const varName : PAnsiChar; var values : PTValued; var labels : PTValuec; var numLabels : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  TspssGetVarNames           = function(const handle : integer; var numVars : integer; var varNames : PTValuec; var varTypes : PTValuei) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função retorna o valor do label apartir do nome curto da variável}
  TspssGetVarCValueLabels    = function(const handle : integer; const varName : PAnsiChar; var values : PTValuec; var labels : PTValuec; var numLabels : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Escrita no arquivo}
  {28/04/2008}

  {Esta função grava um case no o arquivo de dados especificado                }
  {Ela deve ser chamada após ter sido setado os valores das variáveis, através }
  {spssSetValueNumeric e spssSetValueChar.                                     }
  {handle       - Handle do arquivo.                                           }
  TspssCommitCaseRecord      = function(const handle : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função grava o dicionário de dados no arquivo associado                }
  {Antes da escrita do case o dicionário de dados deve esta comitado           }
  {handle       - Handle do arquivo.                                           }
  TspssCommitHeader          = function(const handle : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Manipulação de data e hora}
  {28/04/2008}

  {Esta função converte data para o formato gregoriano                         }
  {day        - Dias do mês                                                    }
  {month      - meses do ano                                                   }
  {year       - ano a partir de 1582                                           }
  {spssdate   - Data no formato gregoriano                                     }
  TspssConvertDate           = function(day, month, year : integer; Var spssDate : Double) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função converte a data a partir de um valor interno para gregoriano    }
  {day        - Dias do mês                                                    }
  {month      - meses do ano                                                   }
  {year       - ano a partir de 1582                                           }
  {spssdate   - Data no formato gregoriano                                     }
  TspssConvertSPSSDate       = function(var day, month, year : integer; spssDate : Double) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função quebra um valor no formato interno SPSS(data) em número de dias }
  {acrescido da hora, minuto, segundo e valores. Note que o segundo valor ?   }
  {armazenado em um double, uma vez que pode ter uma parte fracionária.        }
  {day        - Dias do mês                                                    }
  {hour       - horas do dia                                                   }
  {minute     - Minutos da hora                                                }
  {second     - Segundos do minuto                                             }
  {spsstime   - Data no formato interno SPSS                                   }
  TspssConvertSPSSTime       = function(var day : Longint; var hour, minute : integer; var second : Double; spssTime : double) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função converte o dia, a hora, minuto e segundos para o padrão SPSS    }
  {day        - Dias do mês                                                    }
  {hour       - horas do dia                                                   }
  {minute     - Minutos da hora                                                }
  {second     - Segundos do minuto                                             }
  {spsstime   - Data no formato interno SPSS                                   }
  TspssConvertTime           = function(day : Longint; hour, minute : Integer;  second : Double; Var spssTime : Double) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Adiciona resposta a pesquisa}
  {28/04/2008}

  {Esta função adiciona um simples atributo ao arquivo                         }
  {Se atributo j?existir ele ser?sobre escrito, para ele não sobre escrevido }
  {deve ser passado -1 no atribsub                                             }
  {handle       - Handle do arquivo.                                           }
  {varName      - Nome da variável                                             }
  {attribName   - Nome do atributo                                             }
  {attribSub    - Origem do atributo ou -1                                     }
  {attribText   - 	Texto usado como o valor do atributo                        }
  TspssAddFileAttribute      = function(const handle : integer; const attribName : PAnsiChar; const attribSub  : integer; const attribText : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função adiciona uma multi resposta tipo string ao dicionário do arquivo}
  {handle       - Handle do arquivo.                                           }
  {mrSetName    - Nome da multipla resposta com tamanho máximo de 64 bytes long}
  {               que começe com o sinalizador "$" e tenha um nome válido      }
  {mrSetLabel   - Label(Pergunta) para a multi resposta com o tamanho máximo de}
  {               256 bytes long. Pode ser vazio ou null para indicar que não é}
  {               o label desejado                                             }
  {isDichotomy  - Se não null serão codificadas no dichotomies e zero de outra }
  {               forma                                                        }
  {countedValue - String terminada em null contendo valor contabilizado.       }
  {               Utilizado quando isDichotomy ?diferente de zero e caso deve }
  {               ser 1? characters long e ignorados caso contrário. Pode ser }
  {               nul se isDichotomy igual a zero                              }
  {varNames     - Array terminado em null contendo o nome das variáveis        }
  {               Todas as variáveis devem ter o nome curto(acho que tamanho   }
  {               máximo tem que ser 8).                                       }
  {numVars      - Numero de variáveis da lista (varNames).                     }
  TspssAddMultRespDefC       = function(const handle : integer; const mrSetName : PAnsiChar; const mrSetLabel : PAnsiChar; isDichotomy : integer; const countedValue : PAnsiChar; const varNames : PTValuec; numVars : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função adiciona uma multi resposta a partir da estrututa no dicionario }
  {Ao adicionar na estrutura, definir o nome único e as variáveis devem existir}
  {e ser do tipo adequado - numérica ou string dependendo do qIsNumeric.       }
  {handle       - Handle do arquivo.                                           }
  {pSet         - Ponteiro para estrutura                                      }
  TspssAddMultRespDefExt     = function(const handle : integer; const spssMultRespDef : PspssMultRespDef)  : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função adiciona uma multi resposta tipo numerica ao dicionário         }
  {handle       - Handle do arquivo.                                           }
  {mrSetName    - Nome da multipla resposta com tamanho máximo de 64 bytes long}
  {               que começe com o sinalizador "$" e tenha um nome válido      }
  {mrSetLabel   - Label(Pergunta) para a multi resposta com o tamanho máximo de}
  {               256 bytes long. Pode ser vazio ou null para indicar que não é}
  {               o label desejado                                             }
  {isDichotomy  - Se não null serão codificadas no dichotomies e zero de outra }
  {               forma                                                        }
  {countedValue - String terminada em null contendo valor contabilizado.       }
  {               Utilizado quando isDichotomy ?diferente de zero e caso deve }
  {               ser 1? characters long e ignorados caso contrário. Pode ser }
  {               nul se isDichotomy igual a zero                              }
  {varNames     - Array terminado em null contendo o nome das variáveis        }
  {               Todas as variáveis devem ter o nome curto(acho que tamanho   }
  {               máximo tem que ser 8).                                       }
  {numVars      - Numero de variáveis da lista (varNames).                     }
  TspssAddMultRespDefN       = function(const handle : integer; const mrSetName  : PAnsiChar; const mrSetLabel : PAnsiChar; isDichotomy : integer; countedValue : Longint; const varNames : PTValuec; numVars : Integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função ?similar a sspsaddfileatribute, mas que opera em uma única     }
  {variável da esrutura de atributos.                                          }
  {handle       - Handle do arquivo.                                           }
  {varName      - Nome da variável                                             }
  {attribName   - Nome do atributo                                             }
  {attribSub    - Origem do atributo ou -1                                     }
  {attribText   - 	Texto usado como o valor do atributo                        }
  TspssAddVarAttribute       = function(const handle : integer; const varName    : PAnsiChar; const attribName : PAnsiChar; const attribSub : integer; const attribText : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Desalocação de mémoria}
  {28/04/2008}

  {Esta função desaloca a memória alocada dinamicamente pelas funções          }
  {spssGetFileAttributes. ou spssGetVarAttributes.                              }
  {attribNames  - Ponteiro para o array nomes de atributos                     }
  {attribText   - Ponteiro para array valores de texto                         }
  {nAttributes  - Numero de elementos do array                                 }
  TspssFreeAttributes        = function(attribNames, attribText : PTValuec; const nAttributes : Integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função ?chamada p/ liberar a memoria alocada pela função              }
  {spssGetDateVariables.                                                       }
  {dateInfo     - Array de dados TRENDS                                        }
  TspssFreeDateVariables     = function(Var dateInfo : PTValuelong) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função libera a memoria alocada pela função spssGetMultRespDefs        }
  TspssFreeMultRespDefs      = function(mrespDefs : PAnsiChar)  : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função libera a memoria alocada pela função spssGetMultRespDefByIndex  }
  TspssFreeMultRespDefStruct = function(pSet : PspssMultRespDef) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função libera da memoria dois arrays PTValuec e numero de labels,      }
  {alocados pela função spssGetVarCValueLabels                                 }
  TspssFreeVarCValueLabels   = function(values : PTValuec; labels : PTValuec; numLabels : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função libera a memoria alocada pela função spssGetVariableSets        }
  TspssFreeVariableSets      = function(varSets : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função libera da memoria os arrays PTValued e PTValuec e o numero de   }
  {labels alocados pela função spssGetVarNValueLabels                          }
  TspssFreeVarNValueLabels   = function(values : PTValued; labels : PTValuec; numLabels : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função libera da memoria os arrays PTValuec e PTValuei e o numero de   }
  {variaveis alocados pela função spssGetVarNames                              }
  TspssFreeVarNames          = function(varNames : PTValuec; varTypes : PTValuei; numVars : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}


  {Pega informações sobre o arquivo}
  {25/04/2008}

  {Esta função retorna o tamanho(size) do case associado ao arquivo            }
  {O tamanho(size) do case retornado ser?utilizado em baixo nivel input/output}
  {pelas output procedures spssWholeCaseIn e spssWholeCaseOut                  }
  TspssGetCaseSize           = function(const handle : integer; Var caseSize : Longint) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {***** Não entendi como funciona *****                                       }
  TspssGetCaseWeightVar      = function(const handle : integer; const varName : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função 1 se arquivo compactado e 0 caso nao seja compactado            }
  TspssGetCompression        = function(const handle : integer; var compSwitch : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função reporta informações sobre data TRENDS no arquivo                }
  {Ela aloca em tempo de execucao o array dateInfo com o tamanho numVars, após }
  {utilização do array ?necessário sua liberação da memoria feita pela função }
  {spssFreeDateVariables                                                       }
  TspssGetDateVariables      = function(const handle : integer; var numVars : integer; var dateInfo : PTValuelong) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Função não necessariamente importarte apartir da versão 6 do SPSS           }
  {Se o arquivo não tiver respostas a função ir?rentornar -1 em casecount, mas}
  {se existir resposta ele ir?retornar prescisamente o numero de cases se     }
  {arquivo descompactado e se estiver compactado ir?retornar uma estimativa   }
  {Não pode ser usada se arquivo aberto spssOpenAppend                         }
  TspssGetEstimatedNofCases  = function(const handle : integer; Var caseCount : Longint) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função retorna todos os atributos do arquivo                           }
  {Função aloca a memoria necessário para retornar os nome e seu valores       }
  TspssGetFileAttributes     = function(const handle : integer; var attribNames : PTValuec; var attribText : PTValuec; var nAttributes : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função provem numero da pagina(Handle) do windows relacionado a arquivo}
  TspssGetFileEncoding = function(const handle: integer; var pszEncoding: PAnsiChar):
      integer;

  {Esta função copia o label do arquivo para o buffer(PAnsiChar) id.               }
  {O tamanho do label ?de 64 e terminado em null                              }
  {Assim, o tamanho do buffer deve ser de pelo menos 65                        }
  TspssGetIdString           = function(const handle : integer; var id : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função retorna a codificação da pagina}
  TspssGetInterfaceEncoding  = function() : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função retorna o numero de multi resposta do dicionario                }
  TspssGetMultRespCount      = function(const handle : integer; var nSets : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função obtem os dados de uma multi resposta sentando ela na estrutura  }
  {Após utilização da função mem. deve ser libera c/ spssFreeMultRespDefStruct }
  TspssGetMultRespDefByIndex = function(const handle : integer; const iSet : integer; var ppSet : PspssMultRespDef) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função recupera as definições a partir de arquivo  de dados SPSS       }
  {A memoria alocada por este função deve ser liberada por spssFreeMultRespDefs}
  TspssGetMultRespDefs       = function(const handle : integer; Var mrespDefs : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função reporta o numero de case existentes no arquivo                  }
  TspssGetNumberofCases      = function(const handle : integer; var numofCases : Longint) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função reporta o numero de variables existentes no arquivo             }
  TspssGetNumberofVariables  = function(const handle : integer; Var numVars : Longint) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função reporta informações sobre o arquivo                             }
  {A informação ?constituído por um conjunto de oito int valores copiados     }
  {de registro tipo 7,  subtipo 3, e ?útil principalmente para depuração      }
  {A ordem do array ?a seguinte:                                              }
  {0 - Numero inicial                                                          }
  {1 - Sub numero inicial                                                      }
  {2 - Numero especial                                                         }
  {3 - Codigo da maquina                                                       }
  {4 - Representação de ponto flutuante(casas decimais)                        }
  {5 - Codigo da compressão                                                    }
  {6 - Codigo big / little-endian                                              }
  {7 - Codigo para representação de string                                     }
  TspssGetReleaseInfo        = function(const handle : integer; var relinfo : PTValuei) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função retorna o nome do sistema em que o arquivo foi criado           }
  TspssGetSystemString       = function(const handle : integer; var sysName : array of char) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função coloca os dados do texto criado por TextSmart}
  TspssGetTextInfo           = function(const handle : integer; var textInfo : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função retorna a data de criação do arquivo}
  TspssGetTimeStamp          = function(const handle : integer; var fileDate : PAnsiChar; Var fileTime : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Pegas as informações do case}

  {Esta função pega uma valor do timpo string do case corrente}
  TspssGetValueChar          = function(const handle : integer; varHandle : Double; value : PAnsiChar; valueSize : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função retorna o handle da variavel, que pode ser usado tanto para     }
  {leitura quanto para escrita do valor da variavel                            }
  TspssGetVarHandle          = function(const handle : integer; const varName : PAnsiChar; var varHandle : Double) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função l?o proximo case do arquivo}
  TspssReadCaseRecord        = function(const handle : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função pega uma valor do timpo numeric do case corrente}
  TspssGetValueNumeric       = function(const handle : integer; varHandle : Double ; var value : Double) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função faz a mesma coisa da spssGetFileAttributes. Retorna os atributos de uma variavel}
  TspssGetVarAttributes      = function(const handle : integer; const varName : PAnsiChar; var attribNames : PTValuec; var attribText : PTValuec; var nAttributes : integer)  : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função reporta o valor do atributo de alinhamento da variável}
  TspssGetVarAlignment       = function(const handle : integer; const varName : PAnsiChar; var alignment : Integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função indica se a variavel tipo string possui valor}
  TspssGetVarCMissingValues  = function(const handle : integer; const varName : PAnsiChar; var missingFormat : integer; var missingVal1 : PAnsiChar; Var missingVal2 : PAnsiChar; Var missingVal3 : PAnsiChar)  : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função reporta o tamanho(width) de uma variável}
  TspssGetVarColumnWidth     = function(const handle : integer; const varName : PAnsiChar; var columnWidth :  Integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função reporta o nome curto da variável}
  TspssGetVarCompatName      = function(const handle : integer; const longName : PAnsiChar; var shortName : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função permite ao usuário controlar o numero de caracteres ele deseja  }
  {ler, caso o valor da variavel seja maior que quantidade de caracteres que o }
  {usuario deseja ler ?retornado um erro                                      }
  TspssGetVarCValueLabelLong = function(const handle : integer; const varName : PAnsiChar; const value : PAnsiChar; var labelBuff : PAnsiChar; lenBuff : Integer; var lenLabel : Integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função reporta as informações da variaveis do arquivo de dados         }
  TspssGetVariableSets       = function(const handle : integer; varNames : PTValuec) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função pega nome e o tipo de uma variavel presente no arquivo          }
  TspssGetVarInfo            = function(const handle : integer; const iVar : Integer; var varName : PAnsiChar; var varType : Integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função copia o label da variável passada para função para ponteiro     }
  TspssGetVarLabel           = function(const handle : integer; const varName : PAnsiChar; varLabel : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função retorna o label associado a variável trazendo apenas o numero de}
  {bytes informado                                                             }
  TspssGetVarLabelLong       = function(const handle : integer; const varName : PAnsiChar;  labelBuff : PAnsiChar; lenBuff : integer; var lenLabel : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função retorna o valor medio do level atribuido a variavel             }
  TspssGetVarMeasureLevel    = function(const handle : integer; const varName : PAnsiChar; var measureLevel : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função ?uma combinação para busca de valores de variaveis             }
  TspssGetVarNMissingValues  = function(const handle : integer; const varName : PAnsiChar; var missingFormat : Integer; var missingVal1 : Double; var missingVal2 : Double; var missingVal3 : Double) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função retorna valor de uma label que corresponde a um valor especifico}
  {de uma variavel numerica. Esta permite ao usuario limitar o quantidade de   }
  {byte que deve ser retornado                                                 }
  TspssGetVarNValueLabelLong = function(const handle : integer; const varName : PAnsiChar; value : double; var labelBuff : PAnsiChar; lenBuff : integer; var lenLabel : Integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função retorna formatação(PrintType, Decimal e tamanho) de uma variavel}
  TspssGetVarPrintFormat     = function(const handle : integer; const varName : PAnsiChar; var printType : integer; var printDec : Integer; var printWid : Integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função escreve a formatação de uma variavel no arquivo                 }
  TspssGetVarWriteFormat     = function(const handle : integer; const varName : PAnsiChar; var writeType : Integer; var writeDec : Integer; var writeWid : Integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função determina a se a codificação do arquivo ?compativel            }
  TspssIsCompatibleEncoding  = function (const handle : integer; var bCompatible : Integer):integer ; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta procspssIsCompatibleEncodingedure retorna o menor e o maior valor para uma faixa de valores    }
  {não encontrados                                                             }
  TspssLowHighVal            = procedure (var lowest : Double; var highest : Double); {$IFDEF WIN32} stdcall; {$ENDIF}

  {Retorna o subtipo do arquivo}
  TspssQueryType7            = function(const handle : integer; const subType : integer; var bFound : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Seta o tipo de compressão}
  TspssSetCompression        = function(const handle : integer; compSwitch : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Seta as variáveis TREND}
  TspssSetDateVariables      = function(const handle : integer; numofElements : integer; const dateInfo : Longint) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}




  {Reescreve os atributos do arquivo}
  TspssSetFileAttributes     = function(const handle : integer; const attribNames : TValuec; const attribText : TValuec; const nAttributes : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função seta o label do arquivo de dados SPSS associado ao handle do id informado}
  TspssSetIdString           = function(const handle : integer; const id : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função muda o codigo da interface, se executada com sucesso codigo do arquivo ?alterado}
  TspssSetInterfaceEncoding     = function(const iEncoding : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função retorna a localidade default do arquivo}
  TspssSetLocale             = function(const iCategory : integer; const pszLocale : PAnsiChar) : PAnsiChar; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função ?usada para escrita de multiresposta no arquivo}
  TspssSetMultRespDefs       = function(const handle : integer; const mrespDefs : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Nao entendi direito - mas seta o diretorio temporário}
  TspssSetTempDir            = function(const dirName : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função seta o texto no textinfo do arquivo}
  TspssSetTextInfo           = function(const handle : integer; const textInfo : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função seta o valor em campo string do case corrente}
  TspssSetValueChar          = function(const handle : integer; varHandle : double; const value : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função seta o valor em campo numeric do case corrente}
  TspssSetValueNumeric       = function(const handle : integer; varHandle : double; value : double) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função seta o valor de alinhamento do atributo da variável}
  TspssSetVarAlignment       = function(const handle : integer; const varName : PAnsiChar; alignment : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função reescreve os atributos de uma variável}
  TspssSetVarAttributes      = function(const handle : integer; const varName : PAnsiChar; const attribNames : TValuec; const attribText : TValuec; const nAttributes : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função seta o falta de valor de uma variável string}
  TspssSetVarCMissingValues  = function(const handle : integer; const varName : PAnsiChar; missingFormat : integer; const missingVal1 : PAnsiChar; const missingVal2 : PAnsiChar; const missingVal3 : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta seta o tamanho(width) de uma coluna de uma variável}
  TspssSetVarColumnWidth     = function(const handle : integer; const varName : PAnsiChar; columnWidth : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função altera ou adiciona a um label de uma variável string}
  TspssSetVarCValueLabel     = function(const handle : integer; const varName : PAnsiChar; const value : PAnsiChar; const plabel : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função altera ou adiciona um ou mais labels de uma ou mais variáveis string}
  TspssSetVarCValueLabels    = function(const handle : integer; const varNames : TValuec; numVars : integer; const values : TValuec; const labels : TValuec; numLabels : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função seta o label de uma variável}
  TspssSetVarLabel           = function(const handle : integer; const varName : PAnsiChar; const varLabel : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função seta o valor medio de uma level do atributo da variável}
  TspssSetVarMeasureLevel    = function(const handle : integer; const varName : PAnsiChar; measureLevel : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função seta a falta de valor de uma variável numeric}
  TspssSetVarNMissingValues  = function(const handle : integer; const varName : PAnsiChar; missingFormat : integer; missingVal1 : double; missingVal2 : double; missingVal3 : double) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função altera ou adiciona a um label de uma variável numeric}
  TspssSetVarNValueLabel     = function(const handle : integer; const varName : PAnsiChar; value : double; const plabel : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função altera ou adiciona um ou mais labels de uma ou mais variáveis numeric}
  TspssSetVarNValueLabels    = function(const handle : integer; const varNames : TValuec; numVars : integer; const values : double; const labels : TValuec; numLabels : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função cria uma nova variável string ou numeric no arquivo}
  TspssSetVarName            = function( const handle : integer; const varName : PAnsiChar;  const varLength : integer) : integer;{$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função seta as formatação de impressão da variável}
  TspssSetVarPrintFormat     = function(const handle : integer; const varName : PAnsiChar; printType : integer; printDec : integer; printWid : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função seta as formatação de escrita da variável}
  TspssSetVarWriteFormat     = function(const handle : integer; const varName : PAnsiChar; writeType : integer; writeDec : integer; writeWid : integer) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função define a variável apresenta informações no arquivo de dados.}
  TspssSetVariableSets       = function(const handle : integer; const varSets : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função l?o case do arquivo dentro do ponteiro}
  TspssWholeCaseIn           = function(const handle : integer; var caseRec : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}

  {Esta função escreve o case no arquivo}
  TspssWholeCaseOut          = function(const handle : integer; const caseRec : PAnsiChar) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}


  {Não entendi}
  TspssSeekNextCase          = function(const handle : integer; const caseNumber : Longint) : integer; {$IFDEF WIN32} stdcall; {$ENDIF}


var
  spssOpenRead : TspssOpenRead;
  spssCloseRead : TspssCloseRead;

  spssGetVarNValueLabels : TspssGetVarNValueLabels;
  spssFreeVarNValueLabels : TspssFreeVarNValueLabels;

  spssGetVarNames : TspssGetVarNames;
  spssFreeVarNames : TspssFreeVarNames;

  spssGetVarCValueLabels : TspssGetVarCValueLabels;
  spssFreeVarCValueLabels : TspssFreeVarCValueLabels;

  spssGetMultRespCount : TspssGetMultRespCount;

  spssAddFileAttribute         : TspssAddFileAttribute;
  spssAddMultRespDefC          : TspssAddMultRespDefC;
  spssAddMultRespDefExt        : TspssAddMultRespDefExt;
  spssAddMultRespDefN          : TspssAddMultRespDefN;
  spssAddVarAttribute          : TspssAddVarAttribute;
  spssCloseAppend              : TspssCloseAppend;
  spssCloseWrite               : TspssCloseWrite;
  spssCommitCaseRecord         : TspssCommitCaseRecord;
  spssCommitHeader             : TspssCommitHeader;
  spssConvertDate              : TspssConvertDate;
  spssConvertSPSSDate          : TspssConvertSPSSDate;
  spssConvertSPSSTime          : TspssConvertSPSSTime;
  spssConvertTime              : TspssConvertTime;
  spssCopyDocuments            : TspssCopyDocuments;
  spssFreeAttributes           : TspssFreeAttributes;
  spssFreeDateVariables        : TspssFreeDateVariables;
  spssFreeMultRespDefs         : TspssFreeMultRespDefs;
  spssFreeMultRespDefStruct    : TspssFreeMultRespDefStruct;
  spssFreeVariableSets         : TspssFreeVariableSets;
  spssGetEstimatedNofCases     : TspssGetEstimatedNofCases;
  spssGetFileEncoding          : TspssGetFileEncoding;
  spssGetIdString              : TspssGetIdString;
  spssGetInterfaceEncoding     : TspssGetInterfaceEncoding;
  spssGetMultRespDefByIndex    : TspssGetMultRespDefByIndex;
  spssGetMultRespDefs          : TspssGetMultRespDefs;
  spssGetNumberofCases         : TspssGetNumberofCases;
  spssGetNumberofVariables     : TspssGetNumberofVariables;
  spssGetReleaseInfo           : TspssGetReleaseInfo;
  spssGetSystemString          : TspssGetSystemString;
  spssGetTextInfo              : TspssGetTextInfo;
  spssGetTimeStamp             : TspssGetTimeStamp;
  spssGetValueChar             : TspssGetValueChar;
  spssGetVarHandle             : TspssGetVarHandle;
  spssReadCaseRecord           : TspssReadCaseRecord;
  spssGetValueNumeric          : TspssGetValueNumeric;
  spssGetVarAttributes         : TspssGetVarAttributes;
  spssGetVarAlignment          : TspssGetVarAlignment;
  spssGetVarCMissingValues     : TspssGetVarCMissingValues;
  spssGetVarColumnWidth        : TspssGetVarColumnWidth;
  spssGetVarCompatName         : TspssGetVarCompatName;
  spssGetVariableSets          : TspssGetVariableSets;
  spssGetVarInfo               : TspssGetVarInfo;
  spssGetVarLabel              : TspssGetVarLabel;
  spssGetVarLabelLong          : TspssGetVarLabelLong;
  spssGetVarMeasureLevel       : TspssGetVarMeasureLevel;
  spssGetVarNMissingValues     : TspssGetVarNMissingValues;
  spssGetVarNValueLabelLong    : TspssGetVarNValueLabelLong;
  spssGetVarPrintFormat        : TspssGetVarPrintFormat;
  spssIsCompatibleEncoding     : TspssIsCompatibleEncoding;
  spssLowHighVal               : TspssLowHighVal;
  spssOpenAppend               : TspssOpenAppend;
  spssOpenWrite                : TspssOpenWrite;

  spssQueryType7               : TspssQueryType7;
  spssSeekNextCase             : TspssSeekNextCase;
  spssSetCompression           : TspssSetCompression;
  spssSetDateVariables         : TspssSetDateVariables;
  spssSetFileAttributes        : TspssSetFileAttributes;

  spssSetIdString              : TspssSetIdString;
  spssSetInterfaceEncoding        : TspssSetInterfaceEncoding;
  spssSetMultRespDefs          : TspssSetMultRespDefs;
  spssSetTempDir               : TspssSetTempDir;
  spssSetTextInfo              : TspssSetTextInfo;
  spssSetValueChar             : TspssSetValueChar;
  spssSetValueNumeric          : TspssSetValueNumeric;
  spssSetVarAlignment          : TspssSetVarAlignment;
  spssSetVarAttributes         : TspssSetVarAttributes;
  spssSetVarCMissingValues     : TspssSetVarCMissingValues;
  spssSetVarColumnWidth        : TspssSetVarColumnWidth;
  spssSetVarCValueLabel        : TspssSetVarCValueLabel;
  spssSetVarCValueLabels       : TspssSetVarCValueLabels;
  spssSetVarLabel              : TspssSetVarLabel;
  spssSetVarMeasureLevel       : TspssSetVarMeasureLevel;
  spssSetVarNMissingValues     : TspssSetVarNMissingValues;
  spssSetVarNValueLabel        : TspssSetVarNValueLabel;
  spssSetVarNValueLabels       : TspssSetVarNValueLabels;
  spssSetVarName               : TspssSetVarName;
  spssSetVarPrintFormat        : TspssSetVarPrintFormat;
  spssSetVariableSets          : TspssSetVariableSets;
  spssWholeCaseIn              : TspssWholeCaseIn;
  spssWholeCaseOut             : TspssWholeCaseOut;



  spssGetCaseSize              : TspssGetCaseSize;
  spssGetCaseWeightVar         : TspssGetCaseWeightVar;
  spssGetCompression           : TspssGetCompression;
  spssGetDateVariables         : TspssGetDateVariables;
  spssGetFileAttributes        : TspssGetFileAttributes;
  spssSetLocale                : TspssSetLocale;
  spssGetWriteFormat           : TspssSetVarWriteFormat;

implementation
var
  _handle : THandle = 0;

function SetProcAddress() : Boolean;
begin

  result := false;

  @spssOpenRead := GetProcAddress(_handle, 'spssOpenRead');
  if (@spssOpenRead = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssOpenRead"');
    exit;
  end;

  @spssCloseRead := GetProcAddress(_handle, 'spssCloseRead');
  if (@spssCloseRead = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssCloseRead"');
    exit;
  end;

  @spssGetVarNValueLabels := GetProcAddress(_handle, 'spssGetVarNValueLabels');
  if (@spssGetVarNValueLabels = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetVarNValueLabels"');
    exit;
  end;

  @spssFreeVarNValueLabels := GetProcAddress(_handle, 'spssFreeVarNValueLabels');
  if (@spssFreeVarNValueLabels = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssFreeVarNValueLabels"');
    exit;
  end;

  @spssGetVarNames := GetProcAddress(_handle, 'spssGetVarNames');
  if (@spssGetVarNames = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetVarNames"');
    exit;
  end;

  @spssFreeVarNames := GetProcAddress(_handle, 'spssFreeVarNames');
  if (@spssFreeVarNames = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssFreeVarNames"');
    exit;
  end;

  @spssGetVarCValueLabels := GetProcAddress(_handle, 'spssGetVarCValueLabels');
  if (@spssGetVarCValueLabels = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetVarCValueLabels"');
    exit;
  end;

  @spssFreeVarCValueLabels := GetProcAddress(_handle, 'spssFreeVarCValueLabels');
  if (@spssFreeVarCValueLabels = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssFreeVarCValueLabels"');
    exit;
  end;

  @spssGetMultRespCount := GetProcAddress(_handle, 'spssGetMultRespCount');
  if (@spssGetMultRespCount = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetMultRespCount"');
    exit;
  end;

  @spssAddFileAttribute := GetProcAddress(_handle, 'spssAddFileAttribute');
  if (@spssAddFileAttribute = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssAddFileAttribute"');
    exit;
  end;

  @spssAddMultRespDefC := GetProcAddress(_handle, 'spssAddMultRespDefC');
  if (@spssAddMultRespDefC = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssAddMultRespDefC"');
    exit;
  end;

  @spssAddMultRespDefExt := GetProcAddress(_handle, 'spssAddMultRespDefExt');
  if (@spssAddMultRespDefExt = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssAddMultRespDefExt"');
    exit;
  end;

  @spssAddMultRespDefN := GetProcAddress(_handle, 'spssAddMultRespDefN');
  if (@spssAddMultRespDefN = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssAddMultRespDefN"');
    exit;
  end;

  @spssAddVarAttribute := GetProcAddress(_handle, 'spssAddVarAttribute');
  if (@spssAddVarAttribute = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssAddVarAttribute"');
    exit;
  end;

  @spssCloseAppend := GetProcAddress(_handle, 'spssCloseAppend');
  if (@spssCloseAppend = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssCloseAppend"');
    exit;
  end;

  @spssCloseWrite := GetProcAddress(_handle, 'spssCloseWrite');
  if (@spssCloseWrite = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssCloseWrite"');
    exit;
  end;

  @spssCommitCaseRecord := GetProcAddress(_handle, 'spssCommitCaseRecord');
  if (@spssCommitCaseRecord = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssCommitCaseRecord"');
    exit;
  end;

  @spssConvertDate := GetProcAddress(_handle, 'spssConvertDate');
  if (@spssConvertDate = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssConvertDate"');
    exit;
  end;

  @spssConvertSPSSDate := GetProcAddress(_handle, 'spssConvertSPSSDate');
  if (@spssConvertSPSSDate = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssConvertSPSSDate"');
    exit;
  end;

  @spssConvertSPSSTime := GetProcAddress(_handle, 'spssConvertSPSSTime');
  if (@spssConvertSPSSTime = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssConvertSPSSTime"');
    exit;
  end;

  @spssConvertTime := GetProcAddress(_handle, 'spssConvertTime');
  if (@spssConvertTime = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssConvertTime"');
    exit;
  end;

  @spssCopyDocuments := GetProcAddress(_handle, 'spssCopyDocuments');
  if (@spssCopyDocuments = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssCopyDocuments"');
    exit;
  end;

  @spssFreeAttributes := GetProcAddress(_handle, 'spssFreeAttributes');
  if (@spssFreeAttributes = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssFreeAttributes"');
    exit;
  end;

  @spssFreeDateVariables := GetProcAddress(_handle, 'spssFreeDateVariables');
  if (@spssFreeDateVariables = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssFreeDateVariables"');
    exit;
  end;

  @spssFreeMultRespDefs := GetProcAddress(_handle, 'spssFreeMultRespDefs');
  if (@spssFreeMultRespDefs = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssFreeMultRespDefs"');
    exit;
  end;

  @spssFreeMultRespDefStruct := GetProcAddress(_handle, 'spssFreeMultRespDefStruct');
  if (@spssFreeMultRespDefStruct = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssFreeMultRespDefStruct"');
    exit;
  end;

  @spssFreeVariableSets := GetProcAddress(_handle, 'spssFreeVariableSets');
  if (@spssFreeVariableSets = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssFreeVariableSets"');
    exit;
  end;

  @spssGetEstimatedNofCases := GetProcAddress(_handle, 'spssGetEstimatedNofCases');
  if (@spssGetEstimatedNofCases = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssFreeVariableSets"');
    exit;
  end;

  {@spssGetFileEncoding := GetProcAddress(_handle, 'spssGetFileEncoding');
  if (@spssGetFileEncoding = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetFileEncoding"');
    exit;
  end;}

  @spssGetIdString := GetProcAddress(_handle, 'spssGetIdString');
  if (@spssGetIdString = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetIdString"');
    exit;
  end;

  @spssGetInterfaceEncoding := GetProcAddress(_handle, 'spssGetInterfaceEncoding');
  if (@spssGetInterfaceEncoding = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetInterfaceEncoding"');
    exit;
  end;

  @spssGetMultRespDefByIndex := GetProcAddress(_handle, 'spssGetMultRespDefByIndex');
  if (@spssGetMultRespDefByIndex = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetMultRespDefByIndex"');
    exit;
  end;

  @spssGetMultRespDefs := GetProcAddress(_handle, 'spssGetMultRespDefs');
  if (@spssGetMultRespDefs = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetMultRespDefs"');
    exit;
  end;

  @spssGetNumberofCases := GetProcAddress(_handle, 'spssGetNumberofCases');
  if (@spssGetNumberofCases = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetNumberofCases"');
    exit;
  end;

  @spssGetNumberofVariables := GetProcAddress(_handle, 'spssGetNumberofVariables');
  if (@spssGetNumberofVariables = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetNumberofVariables"');
    exit;
  end;

  @spssGetReleaseInfo := GetProcAddress(_handle, 'spssGetReleaseInfo');
  if (@spssGetReleaseInfo = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetReleaseInfo"');
    exit;
  end;

  @spssGetSystemString := GetProcAddress(_handle, 'spssGetSystemString');
  if (@spssGetSystemString = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetSystemString"');
    exit;
  end;

  @spssGetTextInfo := GetProcAddress(_handle, 'spssGetTextInfo');
  if (@spssGetTextInfo = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetTextInfo"');
    exit;
  end;

  @spssGetTimeStamp := GetProcAddress(_handle, 'spssGetTimeStamp');
  if (@spssGetTimeStamp = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetTimeStamp"');
    exit;
  end;

  @spssGetValueChar := GetProcAddress(_handle, 'spssGetValueChar');
  if (@spssGetValueChar = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetValueChar"');
    exit;
  end;

  @spssGetVarHandle := GetProcAddress(_handle, 'spssGetVarHandle');
  if (@spssGetVarHandle = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetVarHandle"');
    exit;
  end;

  @spssReadCaseRecord := GetProcAddress(_handle, 'spssReadCaseRecord');
  if (@spssReadCaseRecord = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssReadCaseRecord"');
    exit;
  end;

  @spssGetValueNumeric := GetProcAddress(_handle, 'spssGetValueNumeric');
  if (@spssGetValueNumeric = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetValueNumeric"');
    exit;
  end;

  @spssGetVarAttributes := GetProcAddress(_handle, 'spssGetVarAttributes');
  if (@spssGetVarAttributes = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetVarAttributes"');
    exit;
  end;

  @spssGetVarAlignment := GetProcAddress(_handle, 'spssGetVarAlignment');
  if (@spssGetVarAlignment = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetVarAlignment"');
    exit;
  end;

  @spssGetVarCMissingValues := GetProcAddress(_handle, 'spssGetVarCMissingValues');
  if (@spssGetVarCMissingValues = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetVarCMissingValues"');
    exit;
  end;

  @spssGetVarColumnWidth := GetProcAddress(_handle, 'spssGetVarColumnWidth');
  if (@spssGetVarColumnWidth = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetVarColumnWidth"');
    exit;
  end;

  @spssGetVarCompatName := GetProcAddress(_handle, 'spssGetVarCompatName');
  if (@spssGetVarCompatName = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetVarCompatName"');
    exit;
  end;

  @spssGetVariableSets := GetProcAddress(_handle, 'spssGetVariableSets');
  if (@spssGetVariableSets = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetVariableSets"');
    exit;
  end;

  @spssGetVarInfo := GetProcAddress(_handle, 'spssGetVarInfo');
  if (@spssGetVarInfo = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetVarInfo"');
    exit;
  end;

  @spssGetVarLabel := GetProcAddress(_handle, 'spssGetVarLabel');
  if (@spssGetVarLabel = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetVarLabel"');
    exit;
  end;

  @spssGetVarLabelLong := GetProcAddress(_handle, 'spssGetVarLabelLong');
  if (@spssGetVarLabelLong = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetVarLabelLong"');
    exit;
  end;

  @spssGetVarMeasureLevel := GetProcAddress(_handle, 'spssGetVarMeasureLevel');
  if (@spssGetVarMeasureLevel = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetVarMeasureLevel"');
    exit;
  end;

  @spssGetVarNMissingValues := GetProcAddress(_handle, 'spssGetVarNMissingValues');
  if (@spssGetVarNMissingValues = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetVarNMissingValues"');
    exit;
  end;

  @spssGetVarNValueLabelLong := GetProcAddress(_handle, 'spssGetVarNValueLabelLong');
  if (@spssGetVarNValueLabelLong = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetVarNValueLabelLong"');
    exit;
  end;

 {@spssIsCompatibleEncoding := GetProcAddress(_handle, 'spssIsCompatibleEncoding@8');
  if (@spssIsCompatibleEncoding = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssisCompatibleEncoding"');
    exit;
  end; }

  @spssLowHighVal := GetProcAddress(_handle, 'spssLowHighVal');
  if (@spssLowHighVal = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssLowHighVal"');
    exit;
  end;

  @spssOpenAppend := GetProcAddress(_handle, 'spssOpenAppend');
  if (@spssOpenAppend = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssOpenAppend"');
    exit;
  end;

  @spssOpenWrite := GetProcAddress(_handle, 'spssOpenWrite');
  if (@spssOpenAppend = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssWrite"');
    exit;
  end;

  @spssGetVarPrintFormat := GetProcAddress(_handle, 'spssGetVarPrintFormat');
  if (@spssGetVarPrintFormat = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetVarPrintFormat"');
    exit;
  end;

  @spssQueryType7 := GetProcAddress(_handle, 'spssQueryType7');
  if (@spssQueryType7 = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssQueryType7"');
    exit;
  end;

  @spssReadCaseRecord := GetProcAddress(_handle, 'spssReadCaseRecord');
  if (@spssReadCaseRecord = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssReadCaseRecord"');
    exit;
  end;

  @spssSeekNextCase := GetProcAddress(_handle, 'spssSeekNextCase');
  if (@spssSeekNextCase = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSeekNextCase"');
    exit;
  end;

  @spssSetCompression := GetProcAddress(_handle, 'spssSetCompression');
  if (@spssSetCompression = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetCompression"');
    exit;
  end;

  @spssSetDateVariables := GetProcAddress(_handle, 'spssSetDateVariables');
  if (@spssSetDateVariables = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetDateVariables"');
    exit;
  end;

  @spssSetFileAttributes := GetProcAddress(_handle, 'spssSetFileAttributes');
  if (@spssSetFileAttributes = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetFileAttributes"');
    exit;
  end;

  @spssSetIdString := GetProcAddress(_handle, 'spssSetIdString');
  if (@spssSetIdString = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetIdString"');
    exit;
  end;

  @spssSetInterfaceEncoding := GetProcAddress(_handle, 'spssSetInterfaceEncoding');
  if (@spssSetInterfaceEncoding = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssInterfaceEncoding"');
    exit;
  end;

  @spssSetMultRespDefs := GetProcAddress(_handle, 'spssSetMultRespDefs');
  if (@spssSetMultRespDefs = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetMultRespDefs"');
    exit;
  end;

  @spssSetTempDir := GetProcAddress(_handle, 'spssSetTempDir');
  if (@spssSetTempDir = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetTempDir"');
    exit;
  end;

  @spssSetTextInfo := GetProcAddress(_handle, 'spssSetTextInfo');
  if (@spssSetTextInfo = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetTextInfo"');
    exit;
  end;

  @spssSetValueChar := GetProcAddress(_handle, 'spssSetValueChar');
  if (@spssSetValueChar = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetValueChar"');
    exit;
  end;

  @spssSetValueNumeric := GetProcAddress(_handle, 'spssSetValueNumeric');
  if (@spssSetValueNumeric = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetValueNumeric"');
    exit;
  end;

  @spssSetVarAlignment := GetProcAddress(_handle, 'spssSetVarAlignment');
  if (@spssSetVarAlignment = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetVarAlignment"');
    exit;
  end;

  @spssSetVarAttributes := GetProcAddress(_handle, 'spssSetVarAttributes');
  if (@spssSetVarAttributes = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetVarAttributes"');
    exit;
  end;

  @spssSetVarCMissingValues := GetProcAddress(_handle, 'spssSetVarCMissingValues');
  if (@spssSetVarCMissingValues = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetVarCMissingValues"');
    exit;
  end;

  @spssSetVarColumnWidth := GetProcAddress(_handle, 'spssSetVarColumnWidth');
  if (@spssSetVarColumnWidth = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetVarColumnWidth"');
    exit;
  end;

  @spssSetVarCValueLabel := GetProcAddress(_handle, 'spssSetVarCValueLabel');
  if (@spssSetVarCValueLabel = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetVarCValueLabel"');
    exit;
  end;

  @spssSetVarCValueLabels := GetProcAddress(_handle, 'spssSetVarCValueLabels');
  if (@spssSetVarCValueLabels = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetVarCValueLabels"');
    exit;
  end;

  @spssSetVarLabel := GetProcAddress(_handle, 'spssSetVarLabel');
  if (@spssSetVarLabel = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetVarLabel"');
    exit;
  end;

  @spssSetVarMeasureLevel := GetProcAddress(_handle, 'spssSetVarMeasureLevel');
  if (@spssSetVarMeasureLevel = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetVarMeasureLevel"');
    exit;
  end;

  @spssSetVarNMissingValues := GetProcAddress(_handle, 'spssSetVarNMissingValues');
  if (@spssSetVarNMissingValues = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetVarNMissingValues"');
    exit;
  end;

  @spssSetVarNValueLabel := GetProcAddress(_handle, 'spssSetVarNValueLabel');
  if (@spssSetVarNValueLabel = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetVarNValueLabel"');
    exit;
  end;

  @spssSetVarNValueLabels := GetProcAddress(_handle, 'spssSetVarNValueLabels');
  if (@spssSetVarNValueLabels = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetVarNValueLabels"');
    exit;
  end;

  @spssSetVarName := GetProcAddress(_handle, 'spssSetVarName');
  if (@spssSetVarName = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetVarName"');
    exit;
  end;

  @spssSetVarPrintFormat := GetProcAddress(_handle, 'spssSetVarPrintFormat');
  if (@spssSetVarPrintFormat = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetVarPrintFormat"');
    exit;
  end;

  @spssSetVariableSets := GetProcAddress(_handle, 'spssSetVariableSets');
  if (@spssSetVariableSets = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssSetVariableSets"');
    exit;
  end;

  @spssWholeCaseIn := GetProcAddress(_handle, 'spssWholeCaseIn');
  if (@spssWholeCaseIn = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssWholeCaseIn"');
    exit;
  end;

  @spssWholeCaseOut := GetProcAddress(_handle, 'spssWholeCaseOut');
  if (@spssWholeCaseIn = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssWholeCaseOut"');
    exit;
  end;


  @spssGetCaseSize := GetProcAddress(_handle, 'spssGetCaseSize');
  if (@spssGetCaseSize = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetCaseSize"');
    exit;
  end;

  @spssGetCaseWeightVar := GetProcAddress(_handle, 'spssGetCaseWeightVar');
  if (@spssGetCaseWeightVar = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetCaseWeightVar"');
    exit;
  end;

  @spssGetCompression := GetProcAddress(_handle, 'spssGetCompression');
  if (@spssGetCompression = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetCompression"');
    exit;
  end;

  @spssGetDateVariables := GetProcAddress(_handle, 'spssGetDateVariables');
  if (@spssGetDateVariables = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetDateVariables"');
    exit;
  end;

  @spssGetFileAttributes := GetProcAddress(_handle, 'spssGetFileAttributes');
  if (@spssGetFileAttributes = nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssGetFileAttributes"');
    exit;
  end;

  //Eric Add
  @spssCommitHeader :=GetProcAddress(_handle,'spssCommitHeader');
  if (@spssCommitHeader=nil) then begin
    ShowMessage('Impossivel encontrar procedimento "spssCommitHeader"');
    exit;

  end;

  @spssSetLocale :=GetProcAddress(_handle,'spssSetLocale');
  if (@spssSetLocale=nil) then begin

     ShowMessage('Impossive encotrar procedimento "spssSetLocale"');
     exit;


  end;

  result := true;

end;

function LoadLib() : boolean;
begin
  _handle := LoadLibrary(PChar('spssio32.dll'));
  if (_handle <= 0) then begin
      result := false;
      exit;
  end;
  result := true;
end;

function FreeLib() : boolean;
begin
  if (_handle > 0) then begin
    result := FreeLibrary(_handle);
    _handle := 0;
    exit;
  end;
  result := false;
end;

initialization
  if not LoadLib() then begin
    ShowMessage('Não foi possivel carregar SPSSIO32.DLL');
    halt(0);
  end;
  if not SetProcAddress() then
  begin
   halt(0);
  end;

finalization
  FreeLib();

end.
