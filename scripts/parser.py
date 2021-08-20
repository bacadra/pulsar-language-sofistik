import pandas as pd, math
sheet_to_df_map = pd.read_excel('keywords.xlsx', sheet_name=None, header=None)

data = {}

for modulename in sheet_to_df_map:
    module = data[modulename] = {}
    for row in sheet_to_df_map[modulename].itertuples():
        for i, key in enumerate(row):
            if not isinstance(key, str) and math.isnan(key): continue
            if i==0: continue
            if i==1:
                command = module[key] = []
            else:
                command.append(key)


text1 = []
text2 = []

for mname, module in data.items():
    text1.append(rf'''
  {{
  begin: '(?i)^ *([\+-]?prog)( +{mname})( .*)?'
  beginCaptures:
    1: name: 'support.class'
    2: name: 'support.class'
    3: name: 'comment.line'
  patterns: [{{'include': '#{mname}'}}, {{'include': '$self'}}]
  }}
    ''')

    text2.append([rf'''
  {mname}: {{
    patterns: [
      {{include: '#normalText'}}
      {{include: '#normalTextIn'}}
    ''', (command:=[]), '\n    ]\n  }'])

    for cname, commands in module.items():
        command.append(rf'''
      {{
      begin: '(?i)(^ *)({cname})(?=;|$| )'
      beginCaptures:
        2: name: 'keyword.control'
      patterns: [
        {{
          match: '(?i)(?<!\\w)({"|".join(commands)})(?!\\w)'
          name: 'entity.name.function'
        }}
        {{'include': '#normalText'}}
        {{'include': '#normalTextIn'}}
        {{'include': '#{mname}'}}
        {{'include': '$self'}}
        ]
      }}''')



text = (r'''
# ***** References *****
# https://flight-manual.atom.io/hacking-atom/sections/creating-a-grammar/
# http://manual.macromates.com/en/language_grammars/
# https://regex101.com/

scopeName: 'source.sofistik'
name: 'SOFiSTiK'
fileTypes: ['dat','gra','grb','results','sto','dfs','blk']
foldingStartMarker: '+prog'
patterns: [
  {include: '#normalText'}
'''+

''.join(text1)

+r'''

]

repository:

  normalText: {
    patterns: [
      {include: '#defA'}
      {include: '#defC'}
      {include: '#defD'}
      {include: '#inc1'}
      {include: '#comments'}
      {include: '#str1'}
      {include: '#str2'}
      {include: '#linsep'}
      {include: '#var2'}
      {include: '#var3'}
      {include: '#expr'}
      {include: '#units'}
      # {include: '#inLoop'}
    ]
  }
  normalTextIn: {
    patterns: [
      {include: '#var1'}
      {include: '#spec1'}
      {include: '#spec2'}
      {include: '#keyECHO'}
      {include: '#keyUNIT'}
      {include: '#keyPAGE'}
      {include: '#keySIZE'}
      {include: '#keyGNT'}
      {include: '#keyGXX'}
      {include: '#keyGGDP'}
      {include: '#keyGTXT'}
      {include: '#keyGSCA'}
      {include: '#keyGCOL'}
      {include: '#keyGPLI'}
      {include: '#keyGPMI'}
      {include: '#keyGTXI'}
      {include: '#keyGFAI'}
      {include: '#text'}
      {include: '#units'}
    ]
  }

  text: {
    begin: '(?i)^[ ]*(<text>)(?= |$)'
    beginCaptures:
      1: name: 'keyword.control'
    end: '(?i)^[ ]*(<\\/text>)(?= |$)'
    endCaptures:
      1: name: 'keyword.control'
    contentName: 'string'
    patterns: [
        {include: '#var2'}
        {include: '#inc1'}
    ]
  }
  defA: {
    match: '(?i)^[ ]*(#define|#enddef)( .+(?==)| .+|)'
    captures:
      1: name: 'entity.name.section'
      2: name: 'string'
  }
  defC: {
    match: '(?i)^[ ]*(#include|[\\+-]?apply)( +.+)'
    captures:
      1: name: 'entity.name.section'
      2: name: 'string'
  }
  defD: {
    match: '(?i)^[ ]*(#if|#else|#endif)'
    captures:
      1: name: 'entity.name.section'
  }
  inc1: {
    match: '\\$\\(.+?\\)'
    name: 'variable.other'
  }
  comments: {
    match: '(!|//|\\$).*'
    name: 'comment.line'
  }
  str1: { # strings
    match: '\\".*?\\"'
    name: 'string'
  }
  str2: { # strings
    match: "\\'.*?\\'"
    name: 'string'
  }
  linsep: { # line sep
    match: ';'
    name: 'support.type'
  }
  var2: {
    match: '#\\w+'
    name: 'variable.other'
  }
  expr: {
    match: '(?<=\\s|^)(=\\S+)'
    captures:
      1:
        name: 'entity.name.function'
        patterns: [{'include': '#var2'}]
  }
  inLoop: { # what is it? :D
    match: '\\([^ ]+? [^ ]+?(?: [^ ]+?)?\\)'
    name: 'entity.name.function'
  }
  units: {
    match: '(?<=\\w)\\[\\w+?\\]'
    name: 'constant.other'
  }

  var1: {
    match: '(?i)(?<!\\w)(?:let|sto|del|dbg)(?=#\\w+)'
    name: 'keyword.control'
  }
  var3: {
    match: '(?i)(?<!\\w)(?:loop|if|elseif|else|endif)(?!\\w)'
    name: 'keyword.control'
  }
  spec1: {
    match: '(?i)(^ *)(head|txb|txe)( .+?$| *$)'
    captures:
      1: name: 'support.type'
      2: name: 'keyword.control'
  }
  spec2: {
    match: '(?i)(?:^|;)[ ]*(endloop|end)(?!\\w)'
    captures:
      1: name: 'keyword.control'
  }
  keyECHO: {
    match: '(?i)(^ *)(echo)( .+?(?=;|$)| *$)'
    captures:
      1: name: 'support.type'
      2: name: 'keyword.control'
      3: patterns: [
        {'include': '#normalText'}
        {
          match: '(?i)(?<!\\w)(opt|val)(?!\\w)'
          name: 'entity.name.function'
        }
      ]
  }
  keyUNIT: {
    match: '(?i)(^ *)(unit)( .+?(?=;|$)| *$)'
    captures:
      1: name: 'support.type'
      2: name: 'keyword.control'
      3: patterns: [
        {'include': '#normalText'}
        {
          match: '(?i)(?<!\\w)(type|use|dig|set)(?!\\w)'
          name: 'entity.name.function'
        }
      ]
  }
  keyPAGE: {
    match: '(?i)(^ *)(page)( .+?(?=;|$)| *$)'
    captures:
      1: name: 'support.type'
      2: name: 'keyword.control'
      3: patterns: [
        {'include': '#normalText'}
        {
          match: '(?i)(?<!\\w)(firs|line|marg|lano|lani|unio|unii|form|pril|pag)(?!\\w)'
          name: 'entity.name.function'
        }
      ]
  }
  keySIZE: {
    match: '(?i)(^ *)(size)( .+?(?=;|$)| *$)'
    captures:
      1: name: 'support.type'
      2: name: 'keyword.control'
      3: patterns: [
        {'include': '#normalText'}
        {
          match: '(?i)(?<!\\w)(type|sc|w|h|marg|form)(?!\\w)'
          name: 'entity.name.function'
        }
      ]
  }
  keyGNT: {
    match: '(?i)(^ *)(gnt)( .+?(?=;|$)| *$)'
    captures:
      1: name: 'support.type'
      2: name: 'keyword.control'
      3: patterns: [
        {'include': '#normalText'}
        {
          match: '(?i)(?<!\\w)(no|sc|xmin|ymin|xmax|ymax|wxmi|wymi|wxma|wyma)(?!\\w)'
          name: 'entity.name.function'
        }
      ]
  }
  keyGXX: {
    match: '(?i)(^ *)(gnt|gpm|gfa)( .+?(?=;|$)| *$)'
    captures:
      1: name: 'support.type'
      2: name: 'keyword.control'
      3: patterns: [
        {'include': '#normalText'}
        {
          match: '(?i)(?<!\\w)(x1|y1|x1|x2|y2|x2|x3|y3|x3|x4|y4|x4|x5|y5|x5|x6|y6|x6|x7|y7|x7|x8|y8|x8|x9|y9|x9|x10|y10|x10|x11|y11|x11|x12|y12|x12|x13|y13|x13|x14|y14|x14|x15|y15|x15|x16|y16|x16)(?!\\w)'
          name: 'entity.name.function'
        }
      ]
  }
  keyGGDP: {
    match: '(?i)(^ *)(ggdp)( .+?(?=;|$)| *$)'
    captures:
      1: name: 'support.type'
      2: name: 'keyword.control'
      3: patterns: [
        {'include': '#normalText'}
        {
          match: '(?i)(?<!\\w)(type|x1|y1|x1|x2|y2|x2|x3|y3|x3|x4|y4|x4|x5|y5|x5|x6|y6|x6|x7|y7|x7|x8|y8|x8|x9|y9|x9|x10|y10|x10|x11|y11|x11|x12|y12|x12|x13|y13|x13|x14|y14|x14|x15|y15|x15)(?!\\w)'
          name: 'entity.name.function'
        }
      ]
  }
  keyGTXT: {
    match: '(?i)(^ *)(gtxt)( .+?(?=;|$)| *$)'
    captures:
      1: name: 'support.type'
      2: name: 'keyword.control'
      3: patterns: [
        {'include': '#normalText'}
        {
          match: '(?i)(?<!\\w)(x|y|text|val|dim|nd)(?!\\w)'
          name: 'entity.name.function'
        }
      ]
  }
  keyGSCA: {
    match: '(?i)(^ *)(gsca)( .+?(?=;|$)| *$)'
    captures:
      1: name: 'support.type'
      2: name: 'keyword.control'
      3: patterns: [
        {'include': '#normalText'}
        {
          match: '(?i)(?<!\\w)(x1|y1|x2|y2|text|val|dim|nd)(?!\\w)'
          name: 'entity.name.function'
        }
      ]
  }
  keyGCOL: {
    match: '(?i)(^ *)(gcol)( .+?(?=;|$)| *$)'
    captures:
      1: name: 'support.type'
      2: name: 'keyword.control'
      3: patterns: [
        {'include': '#normalText'}
        {
          match: '(?i)(?<!\\w)(col|r|g|b)(?!\\w)'
          name: 'entity.name.function'
        }
      ]
  }
  keyGPLI: {
    match: '(?i)(^ *)(gpli)( .+?(?=;|$)| *$)'
    captures:
      1: name: 'support.type'
      2: name: 'keyword.control'
      3: patterns: [
        {'include': '#normalText'}
        {
          match: '(?i)(?<!\\w)(ind|col|type|widt|scat)(?!\\w)'
          name: 'entity.name.function'
        }
      ]
  }
  keyGPMI: {
    match: '(?i)(^ *)(gpmi)( .+?(?=;|$)| *$)'
    captures:
      1: name: 'support.type'
      2: name: 'keyword.control'
      3: patterns: [
        {'include': '#normalText'}
        {
          match: '(?i)(?<!\\w)(ind|col|type|size)(?!\\w)'
          name: 'entity.name.function'
        }
      ]
  }
  keyGTXI: {
    match: '(?i)(^ *)(gtxi)( .+?(?=;|$)| *$)'
    captures:
      1: name: 'support.type'
      2: name: 'keyword.control'
      3: patterns: [
        {'include': '#normalText'}
        {
          match: '(?i)(?<!\\w)(ind|col|h|bx|by|hali|vali|path|expa|spac|font)(?!\\w)'
          name: 'entity.name.function'
        }
      ]
  }
  keyGFAI: {
    match: '(?i)(^ *)(gfai)( .+?(?=;|$)| *$)'
    captures:
      1: name: 'support.type'
      2: name: 'keyword.control'
      3: patterns: [
        {'include': '#normalText'}
        {
          match: '(?i)(?<!\\w)(ind|col|styl|type)(?!\\w)'
          name: 'entity.name.function'
        }
      ]
  }
'''+

''.join([txt[0]+''.join(txt[1])+txt[2] for txt in text2])
)

with open('../grammars/sofistik.cson', 'w') as f:
    f.write(text.strip())
