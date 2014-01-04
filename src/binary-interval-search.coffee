

#-----------------------------------------------------------------------------------------------------------
module.exports = binary_interval_search = ( data, lo_bound_key, hi_bound_key, id_key, probe ) ->
  ### Given `data`, three indexes, and a `probe`, perform a binary search through `data` to find into which
  of the given intervals `probe` falls.

  `data` must be a list of lists or objects with each entry representing an interval; intervals must not
  overlap, and the intervals in `data` must be sorted in ascending ordered. There may be gaps between
  intervals.

  You must give the keys to the lower and upper boundaries, so that `data[ idx ][ lo_bound_key ]` yields the
  first value and `data[ idx ][ hi_bound_key ]` the last value of each interval. `id_key` should be either
  `null` or a key for an entry so that `data[ idx ][ id_key ]` yields the ID (or whatever info) you want to
  retrieve with your search.

  If a match is found, the result will be either the index of the matching interval, or, if `id_key` was
  defined, `interval[ id_key ]`. If no match is found, `null` is returned.

  With thx to http://googleresearch.blogspot.de/2006/06/extra-extra-read-all-about-it-nearly.html
  ###
  lo_idx    = 0
  hi_idx    = data.length - 1
  #.........................................................................................................
  while lo_idx <= hi_idx
    mid_idx         = Math.floor ( lo_idx + hi_idx ) / 2
    id_and_range    = data[ mid_idx ]
    lo_bound        = id_and_range[ lo_bound_key ]
    hi_bound        = id_and_range[ hi_bound_key ]
    #.......................................................................................................
    if lo_bound <= probe <= hi_bound
      return if id_key? then id_and_range[ id_key ] else mid_idx
    #.......................................................................................................
    if probe < lo_bound then hi_idx = mid_idx - 1 else lo_idx = mid_idx + 1
  #.........................................................................................................
  return null

