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

# function convert_index(str)
#     r = parse(Int, string(collect(c for c in str if isdigit(c))...))
#     c = col(string(collect(c for c in str if !isdigit(c))...))
#     return r, c
# end
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

function readcells(sheet, ss)
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
            ctag = readuntil(io, ">")
            c = ctag[end-1]
            if c != '/'
                rs = search(ctag, "r=\"")
                re = search(ctag, "\"", last(rs) + 1)
                pos = ctag[last(rs) + 1:first(re) - 1]
                str = readuntil(io, "</c>")[1:end-4]
                # str = replace(str, "<v>", "")
                # str = replace(str, "</v>", "")
                str = rm_div(str)
                if ismatch(r"(t=\"s\")",ctag)
                    val = ss[parse(Int, str)]
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
    tdir = joinpath(tempdir(), tempname())
    run(`unzip $file -d $tdir`)
    ss = get_shared_strings(joinpath(tdir, "xl", "sharedStrings.xml"))
    sheets = []
    for f in readdir(joinpath(tdir, "xl", "worksheets"))
        push!(sheets, readcells(joinpath(tdir, "xl", "worksheets", f), ss))
    end
    rm(tdir,recursive = true)
    return sheets
end

export readxls

end
