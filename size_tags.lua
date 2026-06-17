if FORMAT == "pptx" then
  function Blocks(blocks)
    local new_blocks = {}
    for i, el in ipairs(blocks) do
      table.insert(new_blocks, el)
      
      -- Check if the current block is a Header
      if el.t == "Header" then
        -- Inject a speaker note containing the requested size tag
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
