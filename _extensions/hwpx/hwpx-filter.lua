-- hwpx-filter.lua
-- Pandoc Lua filter: serialize AST to JSON, invoke Python to produce .hwpx

function Pandoc(doc)
  -- Serialize entire AST to JSON
  local json_ast = pandoc.write(doc, 'json')

  -- Determine output path: replace .docx with .hwpx
  local output_file = PANDOC_STATE.output_file
  if not output_file or output_file == "" then
    io.stderr:write("[hwpx] WARNING: no output_file detected, skipping HWPX generation\n")
    return doc
  end

  local hwpx_path = output_file:gsub("%.docx$", ".hwpx")

  -- Locate Python script (same directory as this Lua filter)
  local script_dir = PANDOC_SCRIPT_FILE:match("(.*[/\\])")
  if not script_dir then
    script_dir = "./"
  end
  local script_path = script_dir .. "hwpx_writer.py"

  io.stderr:write("[hwpx] Generating " .. hwpx_path .. " ...\n")

  -- Call Python with JSON AST on stdin
  local ok, result_or_err = pcall(function()
    return pandoc.pipe('python3', {script_path, '--output', hwpx_path}, json_ast)
  end)

  if ok then
    io.stderr:write("[hwpx] Successfully created " .. hwpx_path .. "\n")
    -- Schedule .docx cleanup: write a marker file so post-render can delete it
    local marker_path = output_file .. ".hwpx-cleanup"
    local marker = io.open(marker_path, "w")
    if marker then
      marker:write(output_file)
      marker:close()
    end
  else
    io.stderr:write("[hwpx] ERROR: Python script failed: " .. tostring(result_or_err) .. "\n")
    io.stderr:write("[hwpx] The .docx file is preserved at " .. output_file .. "\n")
  end

  return doc
end
