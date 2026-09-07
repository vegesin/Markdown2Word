function Math(el)
  local formula = el.text:gsub("%s+", " "):gsub("^%s+", ""):gsub("%s+$", "")
  if el.mathtype == "DisplayMath" then
    return pandoc.Str("\\[" .. formula .. "\\]")
  end
  return pandoc.Str("$" .. formula .. "$")
end
