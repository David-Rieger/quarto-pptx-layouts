if FORMAT == "pptx" then
  function Blocks(blocks)
    local new_blocks = {}
    for i, el in ipairs(blocks) do
      table.insert(new_blocks, el)
      if el.t == "Header" then
        if el.classes:includes('smaller') then
          local note = pandoc.Div(pandoc.Para(pandoc.Str("[size: smaller]")), pandoc.Attr("", {"notes"}))
          table.insert(new_blocks, note)
        elseif el.classes:includes('smallest') then
          local note = pandoc.Div(pandoc.Para(pandoc.Str("[size: smallest]")), pandoc.Attr("", {"notes"}))
          table.insert(new_blocks, note)
        end
      end
    end
    return new_blocks
  end
end
