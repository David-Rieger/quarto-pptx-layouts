if FORMAT == "pptx" then
  function Blocks(blocks)
    local new_blocks = {}
    local pending_note = nil
    
    for i, el in ipairs(blocks) do
      -- If we hit a new Header, insert any pending note from the PREVIOUS slide first
      if el.t == "Header" and pending_note then
        table.insert(new_blocks, pending_note)
        pending_note = nil
      end
      
      table.insert(new_blocks, el)
      
      -- Check if the current block is a Header and needs a size tag
      if el.t == "Header" then
        if el.classes:includes('smaller') then
          pending_note = pandoc.Div(pandoc.Para(pandoc.Str("[size: smaller]")), pandoc.Attr("", {"notes"}))
        elseif el.classes:includes('smallest') then
          pending_note = pandoc.Div(pandoc.Para(pandoc.Str("[size: smallest]")), pandoc.Attr("", {"notes"}))
        end
      end
    end
    
    -- Insert the last pending note if the document ends
    if pending_note then
      table.insert(new_blocks, pending_note)
    end
    
    return new_blocks
  end
end
