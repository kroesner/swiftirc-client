Attribute VB_Name = "modNumerics"
Option Explicit

Public Const RPL_AWAY = 301
Public Const RPL_WHOISREGNICK = 307
Public Const RPL_WHOISHELPOP = 310
Public Const RPL_WHOISUSER = 311
Public Const RPL_WHOISSERVER = 312
Public Const RPL_WHOISOPERATOR = 313
Public Const RPL_WHOISIDLE = 317
Public Const RPL_ENDOFWHOIS = 318
Public Const RPL_WHOISCHANNELS = 319

Public Const RPL_LISTSTART = 321
Public Const RPL_LIST = 322
Public Const RPL_LISTEND = 323

Public Const RPL_CHANNELMODEIS = 324
Public Const RPL_TOPIC As Integer = 332
Public Const RPL_TOPICWHOTIME As Integer = 333

Public Const RPL_WHOISBOT As Integer = 335

Public Const RPL_INVEXLIST As Integer = 346
Public Const RPL_ENDOFINVEXLIST As Integer = 347
Public Const RPL_EXLIST As Integer = 348
Public Const RPL_ENDOFEXLIST As Integer = 349
Public Const RPL_BANLIST As Integer = 367
Public Const RPL_ENDOFBANLIST As Integer = 368

Public Const RPL_WHOISHOST As Integer = 378

Public Const RPL_QLIST As Integer = 386
Public Const RPL_ENDOFQLIST As Integer = 387
Public Const RPL_ALIST As Integer = 388
Public Const RPL_ENDOFALIST As Integer = 389

Public Const ERR_NOSUCHNICK As Integer = 401

Public Const ERR_TOOMANYCHANNELS As Integer = 405

Public Const ERR_UNKNOWNCOMMAND As Integer = 421

Public Const ERR_ERRONEUSNICKNAME As Integer = 432
Public Const ERR_NICKNAMEINUSE As Integer = 433

Public Const ERR_LINKCHANNEL As Integer = 470
Public Const ERR_CHANNELISFULL As Integer = 471
Public Const ERR_INVITEONLYCHAN As Integer = 473
Public Const ERR_BANNEDFROMCHAN As Integer = 474
Public Const ERR_BADCHANNELKEY As Integer = 475
Public Const ERR_NEEDREGGEDNICK  As Integer = 477

Public Const ERR_NOPRIVILEGES As Integer = 481
Public Const ERR_SECUREONLYCHAN As Integer = 489

Public Const ERR_TOOMANYJOINS As Integer = 500

Public Const ERR_ADMONLY As Integer = 519
Public Const ERR_OPERONLY As Integer = 520


Public Const RPL_HASFILTER As Integer = 536
Public Const RPL_HASFILTEROVER As Integer = 537

Public Const RPL_WHOISSECURE As Integer = 671


' 10/nov/2011
Public Const RPL_WHOLIST As Integer = 352
Public Const RPL_ENDOFWHOLIST As Integer = 315

Public Const RPL_WHOWASHOST As Integer = 314
Public Const RPL_WHOWASUNKNOWN As Integer = 406
Public Const RPL_ENDOFWHOWAS As Integer = 369
