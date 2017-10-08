module xlsxReader

function get_shared_strings(file)
    sharedstrings = readstring(file)
    i = 1:1
    cnt = -1
    res = Dict{Int,String}()
    while true
        cnt += 1
        i1 = search(sharedstrings, r"<si>", last(i))
        i1 == 0:-1 && break
        i = search(sharedstrings, r"<\/si>", last(i1))
        str = sharedstrings[last(i1) + 1:first(i) - 1]
        str = replace(str, "<t>", "")
        str = replace(str, "</t>", "")
        res[cnt] = str
        i == 0:-1 && break
    end
    return res
end

function parse_digits(R)
    n = length(R)
    x = 0
    for (i,c) in enumerate(R)
        d = c - 0x30
        x += Int(d)*10^(n-i)
    end
    x
end

function convert_index(str)
    C = UInt8[]
    R = UInt8[]
    for c in str
        if isdigit(c)
            push!(R, c)
        else
            push!(C, c)
        end
    end
    return parse_digits(R), col(String(C))
end

function get_sheet_range(io)
    seekstart(io)
    readuntil(io, "<dimension")
    readuntil(io, "ref=")
    str = readuntil(io, "/>")[1:end-2]
    str = replace(str, "\"", "")
    strs = split(str, ":")
    rows = range([parse(string(collect(c for c in s if isdigit(c))...)) for s in strs]...)
    cols = range(col.([string(collect(c for c in s if !isdigit(c))...) for s in strs])...)
    return rows, cols
end

# convert column index to numbers
function col(C::String)
    x = 0
    j = 0
    for i = length(C):-1:1
        x += (UInt8(C[i]) - 0x40)*26^j
        j += 1
    end
    x
end

function rm_div(str)
    if startswith(str, "<v>")
        str = str[4:end]
    end
    if endswith(str, "</v>")
        str = str[1:end-4]
    end
    str
end

function seekuntil(io, str, i = 1)
    while !eof(io)
        c = read(io, Char)
        if c == str[i]
            if i == length(str)
                return position(io)
            else
                return seekuntil(io, str, i + 1)
            end
        else
            i = 1
        end
    end
    return -1
end

function read_ctag(io, lc)
    r = ""
    isstring = false
    while !eof(io)
        c = read(io, Char)
        if c == 'r'
            c = read(io, Char)
            if c == '='
                c = read(io, Char)
                if c == '\"'
                    r = readuntil(io, '\"')[1:end-1]
                end
            end
        end
        if c == 't'
            c = read(io, Char)
            if c == '='
                c = read(io, Char)
                if c == '\"'
                    c = read(io, Char)
                    if c == 's'
                        isstring = true
                    end
                end
            end
        end
        if c == '>'
            break
        end
        lc = c
    end
    return r, isstring, lc == '/'
end

function read_val(io)
    val = ""
    while !eof(io)
        c = read(io, Char)
        if c == '<'
            c = read(io, Char)
            if c == 'v'
                c = read(io, Char)
                if c == '>'
                    val = readuntil(io, '<')[1:end-1]
                    break
                end
            end
        end
    end
    return val
end

function readcells(sheet, ss)
    info("Reading $sheet")
    io = open(sheet)
    sizes = get_sheet_range(io)
    out = Array{Any}(last.(sizes)...)
    seekstart(io)
    s1 = seekuntil(io, "<sheetData>")
    s2 = 100000
    seek(io, s1)
    while !eof(io)
        seekuntil(io, "<c") == -1 && break
        c = read(io, Char)
        if c == ' '
            pos, isstring, fslash = read_ctag(io, c)
            if !fslash
                str = read_val(io)
                if isstring
                    val = ss[parse_digits(str)]
                else
                    val = parse(Float64, str)
                end 
                id1 = convert_index(pos)
                out[id1...] = val
            end
        end
    end
    close(io)
    return out
end

function check_types(arr)
    ts = []
    for i = 1:size(arr, 2)
        t = Set{DataType}()
        for j = 2:size(arr,1)
            push!(t, typeof(arr[j, i]))
        end
        push!(ts, collect(keys(t.dict)))
    end
    ts
end

function readxls(file)
    io = open(file)
    if read(io, UInt8, 4) != [0x50, 0x4b, 0x03,0x04] 
        close(io)
        return error("$file is not an Excel file")
    end
    close(io)
    tdir = joinpath(tempdir(), tempname())
    run(`unzip $file -d $tdir`)
    ss = get_shared_strings(joinpath(tdir, "xl", "sharedStrings.xml"))
    sheets = []
    for f in readdir(joinpath(tdir, "xl", "worksheets"))
        !endswith(f, ".xml") && continue
        push!(sheets, readcells(joinpath(tdir, "xl", "worksheets", f), ss))
    end
    rm(tdir,recursive = true)
    return sheets
end

export readxls

end
