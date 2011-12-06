library(codetools)

lfind =
function(what, env = NULL)
{
   w = find(what)
   if(length(w) == 0 && !is.null(env) && exists(what, env))
     w = "namespace"
   w     
}

findMissing =
function(fun, id = deparse(substitute(fun)), env = NULL)
{
   if(is.character(substitute(fun))) {
     id = fun
     fun = get(fun)
   }
   
   globals = findGlobals(fun, merge = FALSE)

  
   where.funs = structure(if(is.null(env)) lapply(globals$functions, find) else lapply(globals$functions, lfind, env), names = globals$functions)
   where.vars = structure(if(is.null(env)) lapply(globals$variables, find) else lapply(globals$variables, lfind, env) , names = globals$variables)

   structure(list(funs = names(where.funs)[sapply(where.funs, length) == 0],
                  vars = names(where.vars)[sapply(where.funs, length) == 0]),
             class = "GlobalVariablesReport")
}

checkNamespace =
function(ns)
{
   if(is.character(ns)) {
    # if(length(grep("package:", ns)) == 0)
    #   ns = paste("package", ns, sep = ":")
     ns = getNamespace(ns)
   }
   obs = objects(ns)
   obs = grep("^\\.__", obs, invert = TRUE, value = TRUE)

   sapply(obs, function(id) {
                  val = get(id, ns)
                  if(is.function(val))
                    findMissing(val, id, ns)
                  else
                    list(funs = character(), vars = character())
               })
          
}

if(FALSE) {
library(RWordXML); library(XML); library(Rcompression); library(ROOXML)

findMissing("wordDynDoc")

obs = objects("package:RWordXML")
m = lapply(obs, findMissing)
sapply(m, function(x) length(x[[1]]) + length(x[[2]]))
}
